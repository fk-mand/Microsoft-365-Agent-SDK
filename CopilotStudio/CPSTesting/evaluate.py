"""
Copilot Studio Agent Evaluation Utility

Reads a CSV file with 'query' and 'expected_answer' columns, sends each query
to a Copilot Studio agent, and evaluates the response against the expected answer.

Usage:
    python evaluate.py input.csv [--output results.csv] [--verbose]
    python evaluate.py input.csv --scorer llm          # use Azure OpenAI as judge
    python evaluate.py input.csv --scorer both         # run both scorers

Environment variables required in .env:
    COPILOTSTUDIOAGENT__AGENTAPPID
    COPILOTSTUDIOAGENT__CLIENTSECRET
    COPILOTSTUDIOAGENT__TENANTID
    COPILOTSTUDIOAGENT__ENVIRONMENTID
    COPILOTSTUDIOAGENT__SCHEMANAME

    # Required only when --scorer is 'llm' or 'both':
    AZURE_OPENAI_ENDPOINT=https://your-resource.openai.azure.com/
    AZURE_OPENAI_API_KEY=your-api-key
    AZURE_OPENAI_DEPLOYMENT=gpt-4o        # or your deployment name
    AZURE_OPENAI_API_VERSION=2024-12-01-preview  # optional, has default
"""

import argparse
import asyncio
import csv
import http.server
import json
import logging
import sys
import threading
import time
import uuid
import webbrowser
from dataclasses import dataclass, field
from datetime import datetime
from difflib import SequenceMatcher
from os import environ
from pathlib import Path
from urllib.parse import urlparse, parse_qs

from dotenv import load_dotenv
from msal import ConfidentialClientApplication
from microsoft_agents.activity import ActivityTypes
from microsoft_agents.copilotstudio.client import ConnectionSettings, CopilotClient
from openai import AzureOpenAI

load_dotenv()

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(message)s",
    datefmt="%H:%M:%S",
)
logger = logging.getLogger(__name__)

# ── Configuration ────────────────────────────────────────────────────────────

REQUIRED_ENV_VARS = [
    "COPILOTSTUDIOAGENT__AGENTAPPID",
    "COPILOTSTUDIOAGENT__CLIENTSECRET",
    "COPILOTSTUDIOAGENT__TENANTID",
    "COPILOTSTUDIOAGENT__ENVIRONMENTID",
    "COPILOTSTUDIOAGENT__SCHEMANAME",
]

SCOPES = ["https://api.powerplatform.com/.default"]


# ── Data classes ─────────────────────────────────────────────────────────────

@dataclass
class TestCase:
    row_number: int
    query: str
    expected_answer: str


@dataclass
class TestResult:
    row_number: int
    query: str
    expected_answer: str
    actual_answer: str
    passed: str  # "PASS", "FAIL", or "N/A" (no expected answer)
    confidence_score: float  # 0.0 – 1.0 (text-similarity)
    latency_seconds: float
    error: str = ""
    llm_score: float = -1.0  # -1 means not evaluated; 0.0–1.0 when scored
    llm_reasoning: str = ""
    llm_passed: str = ""  # "PASS", "FAIL", "N/A", "ERROR", or "" if not used


@dataclass
class EvalSummary:
    total: int = 0
    evaluated: int = 0  # had expected answers
    passed: int = 0
    failed: int = 0
    skipped: int = 0  # no expected answer
    errors: int = 0
    avg_confidence: float = 0.0
    avg_latency: float = 0.0


# ── Authentication (delegated / user login) ──────────────────────────────────

REDIRECT_PORT = 8400
REDIRECT_URI = f"http://localhost:{REDIRECT_PORT}/auth/callback"


class _AuthCallbackHandler(http.server.BaseHTTPRequestHandler):
    """Tiny HTTP handler that captures the OAuth authorization code."""

    auth_code: str | None = None
    state: str | None = None
    error: str | None = None

    def do_GET(self):
        parsed = urlparse(self.path)
        params = parse_qs(parsed.query)

        if "error" in params:
            _AuthCallbackHandler.error = params["error"][0]
            body = (
                "<html><body><h2>Authentication failed</h2>"
                f"<p>{params.get('error_description', [''])[0]}</p>"
                "<p>You can close this tab.</p></body></html>"
            )
        elif "code" in params:
            _AuthCallbackHandler.auth_code = params["code"][0]
            _AuthCallbackHandler.state = params.get("state", [None])[0]
            body = (
                "<html><body><h2>Authentication successful</h2>"
                "<p>You can close this tab and return to the terminal.</p></body></html>"
            )
        else:
            body = "<html><body><p>Unexpected request.</p></body></html>"

        self.send_response(200)
        self.send_header("Content-Type", "text/html")
        self.end_headers()
        self.wfile.write(body.encode())

    def log_message(self, format, *args):
        pass  # suppress default HTTP logs


def acquire_token() -> str:
    """Acquire a delegated access token via browser-based OAuth login."""
    client_id = environ["COPILOTSTUDIOAGENT__AGENTAPPID"]
    client_secret = environ["COPILOTSTUDIOAGENT__CLIENTSECRET"]
    tenant_id = environ["COPILOTSTUDIOAGENT__TENANTID"]
    authority = f"https://login.microsoftonline.com/{tenant_id}"

    msal_app = ConfidentialClientApplication(
        client_id=client_id,
        client_credential=client_secret,
        authority=authority,
    )

    # Try to get a cached token first (silent auth)
    accounts = msal_app.get_accounts()
    if accounts:
        result = msal_app.acquire_token_silent(SCOPES, account=accounts[0])
        if result and "access_token" in result:
            logger.info("Access token acquired from cache")
            return result["access_token"]

    # Interactive browser login
    state = str(uuid.uuid4())
    auth_url = msal_app.get_authorization_request_url(
        scopes=SCOPES,
        state=state,
        redirect_uri=REDIRECT_URI,
    )

    # Start local server to capture the callback
    server = http.server.HTTPServer(("localhost", REDIRECT_PORT), _AuthCallbackHandler)
    server.timeout = 120  # 2 minute timeout for user to log in

    print(f"\n  Opening browser for login...")
    print(f"  (If browser doesn't open, visit the URL below)")
    print(f"  {auth_url[:120]}...\n")
    webbrowser.open(auth_url)

    # Wait for the callback
    _AuthCallbackHandler.auth_code = None
    _AuthCallbackHandler.error = None
    while _AuthCallbackHandler.auth_code is None and _AuthCallbackHandler.error is None:
        server.handle_request()

    server.server_close()

    if _AuthCallbackHandler.error:
        raise RuntimeError(f"Authentication failed: {_AuthCallbackHandler.error}")

    if _AuthCallbackHandler.state != state:
        raise RuntimeError("Invalid OAuth state — possible CSRF attack")

    # Exchange auth code for token
    result = msal_app.acquire_token_by_authorization_code(
        _AuthCallbackHandler.auth_code,
        scopes=SCOPES,
        redirect_uri=REDIRECT_URI,
    )

    if "access_token" not in result:
        error = result.get("error_description", result.get("error", "Unknown error"))
        raise RuntimeError(f"Failed to acquire token: {error}")

    logger.info("Access token acquired via delegated login")
    return result["access_token"]


# ── Copilot client helpers ───────────────────────────────────────────────────

def create_copilot_client(access_token: str) -> CopilotClient:
    settings = ConnectionSettings(
        environment_id=environ["COPILOTSTUDIOAGENT__ENVIRONMENTID"],
        agent_identifier=environ["COPILOTSTUDIOAGENT__SCHEMANAME"],
        cloud=None,
        copilot_agent_type=None,
        custom_power_platform_cloud=None,
    )
    return CopilotClient(settings, access_token)


async def start_conversation(client: CopilotClient) -> str:
    """Start a conversation and return the conversation_id."""
    conversation_id = None
    activities = client.start_conversation(True)
    async for activity in activities:
        if hasattr(activity, "conversation") and activity.conversation:
            conversation_id = activity.conversation.id
    if not conversation_id:
        raise RuntimeError("Failed to obtain conversation_id from start_conversation")
    return conversation_id


async def ask_question(client: CopilotClient, query: str, conversation_id: str) -> str:
    """Send a query and collect the full text response."""
    replies = client.ask_question(query, conversation_id)
    texts = []
    async for reply in replies:
        if reply.type == ActivityTypes.message and reply.text:
            cleaned = reply.text.strip()
            if cleaned.lower() != "processing":
                texts.append(cleaned)
    return "\n".join(texts)


# ── Answer comparison / scoring ──────────────────────────────────────────────

def _normalize(text: str) -> str:
    """Lowercase and collapse whitespace for comparison."""
    return " ".join(text.lower().split())


def compute_confidence(expected: str, actual: str) -> float:
    """
    Compute a confidence score (0.0–1.0) between expected and actual answers.

    Uses a blend of:
      - SequenceMatcher ratio (overall text similarity)
      - Keyword overlap (important terms from expected answer found in actual)
    """
    norm_expected = _normalize(expected)
    norm_actual = _normalize(actual)

    if not norm_expected or not norm_actual:
        return 0.0

    # Exact match
    if norm_expected == norm_actual:
        return 1.0

    # Containment check — if expected is fully contained in actual
    if norm_expected in norm_actual:
        return 0.95

    # Sequence similarity
    seq_ratio = SequenceMatcher(None, norm_expected, norm_actual).ratio()

    # Keyword overlap — words of 3+ chars from expected found in actual
    expected_keywords = {w for w in norm_expected.split() if len(w) >= 3}
    if expected_keywords:
        actual_words = set(norm_actual.split())
        keyword_hits = sum(1 for kw in expected_keywords if kw in actual_words)
        keyword_ratio = keyword_hits / len(expected_keywords)
    else:
        keyword_ratio = seq_ratio

    # Weighted blend: 60% sequence similarity, 40% keyword overlap
    confidence = 0.6 * seq_ratio + 0.4 * keyword_ratio
    return round(min(confidence, 1.0), 4)


PASS_THRESHOLD = 0.6  # confidence >= 0.6 counts as PASS


# ── LLM-as-Judge scorer (Azure OpenAI) ──────────────────────────────────────

LLM_JUDGE_SYSTEM_PROMPT = """\
You are an evaluation judge. You will be given a user query, an expected answer, \
and an actual answer from an AI agent. Your job is to determine how well the \
actual answer matches the expected answer in terms of correctness and completeness.

Respond ONLY with valid JSON (no markdown fences) in this exact format:
{
  "score": <float between 0.0 and 1.0>,
  "reasoning": "<brief 1-2 sentence explanation>"
}

Scoring guide:
- 1.0: The actual answer fully conveys the same information as the expected answer.
- 0.8-0.9: Substantially correct, minor details differ or extra info is included.
- 0.5-0.7: Partially correct, covers some key points but misses others.
- 0.2-0.4: Marginally related but largely incorrect or incomplete.
- 0.0-0.1: Completely wrong or irrelevant.

Focus on semantic meaning, not exact wording. The actual answer may be longer, \
use different phrasing, or include additional context — that is fine as long as \
it conveys the expected information correctly."""


def _get_aoai_client() -> AzureOpenAI:
    """Create an Azure OpenAI client from environment variables."""
    endpoint = environ.get("AZURE_OPENAI_ENDPOINT")
    api_key = environ.get("AZURE_OPENAI_API_KEY")
    api_version = environ.get("AZURE_OPENAI_API_VERSION", "2024-12-01-preview")

    if not endpoint or not api_key:
        raise RuntimeError(
            "Azure OpenAI not configured. Set AZURE_OPENAI_ENDPOINT and "
            "AZURE_OPENAI_API_KEY in your .env file."
        )

    return AzureOpenAI(
        azure_endpoint=endpoint,
        api_key=api_key,
        api_version=api_version,
    )


def llm_judge_score(
    query: str, expected: str, actual: str, aoai_client: AzureOpenAI
) -> tuple[float, str]:
    """
    Use Azure OpenAI to score how well actual matches expected.
    Returns (score, reasoning).
    """
    deployment = environ.get("AZURE_OPENAI_DEPLOYMENT", "gpt-4o")

    user_prompt = (
        f"**User Query:** {query}\n\n"
        f"**Expected Answer:** {expected}\n\n"
        f"**Actual Answer:** {actual}"
    )

    response = aoai_client.chat.completions.create(
        model=deployment,
        messages=[
            {"role": "system", "content": LLM_JUDGE_SYSTEM_PROMPT},
            {"role": "user", "content": user_prompt},
        ],
        temperature=0.0,
        max_tokens=300,
    )

    raw = response.choices[0].message.content.strip()

    # Parse JSON response
    try:
        parsed = json.loads(raw)
        score = float(parsed["score"])
        score = max(0.0, min(1.0, score))
        reasoning = parsed.get("reasoning", "")
    except (json.JSONDecodeError, KeyError, ValueError):
        logger.warning(f"Failed to parse LLM judge response: {raw}")
        score = 0.0
        reasoning = f"Parse error: {raw[:200]}"

    return score, reasoning


# ── CSV I/O ──────────────────────────────────────────────────────────────────

def load_test_cases(csv_path: str) -> list[TestCase]:
    path = Path(csv_path)
    if not path.exists():
        raise FileNotFoundError(f"CSV file not found: {csv_path}")

    cases = []
    with open(path, newline="", encoding="utf-8-sig") as f:
        reader = csv.DictReader(f)

        # Validate columns
        headers = [h.strip().lower() for h in (reader.fieldnames or [])]
        if "query" not in headers:
            raise ValueError(
                f"CSV must have a 'query' column. Found columns: {reader.fieldnames}"
            )

        for i, row in enumerate(reader, start=2):  # row 1 is header
            # Normalize keys to lowercase
            row_lower = {k.strip().lower(): v for k, v in row.items()}
            query = (row_lower.get("query") or "").strip()
            expected = (row_lower.get("expected_answer") or row_lower.get("expected answer") or "").strip()

            if not query:
                logger.warning(f"Row {i}: empty query, skipping")
                continue

            cases.append(TestCase(row_number=i, query=query, expected_answer=expected))

    logger.info(f"Loaded {len(cases)} test case(s) from {csv_path}")
    return cases


def write_results(results: list[TestResult], output_path: str, scorer: str):
    fieldnames = [
        "row_number",
        "query",
        "expected_answer",
        "actual_answer",
    ]
    if scorer in ("text", "both"):
        fieldnames += ["text_passed", "text_confidence"]
    if scorer in ("llm", "both"):
        fieldnames += ["llm_passed", "llm_score", "llm_reasoning"]
    fieldnames += ["latency_seconds", "error"]

    with open(output_path, "w", newline="", encoding="utf-8") as f:
        writer = csv.DictWriter(f, fieldnames=fieldnames)
        writer.writeheader()
        for r in results:
            row = {
                "row_number": r.row_number,
                "query": r.query,
                "expected_answer": r.expected_answer,
                "actual_answer": r.actual_answer,
                "latency_seconds": f"{r.latency_seconds:.2f}",
                "error": r.error,
            }
            if scorer in ("text", "both"):
                row["text_passed"] = r.passed
                row["text_confidence"] = f"{r.confidence_score:.4f}"
            if scorer in ("llm", "both"):
                row["llm_passed"] = r.llm_passed
                row["llm_score"] = f"{r.llm_score:.4f}" if r.llm_score >= 0 else ""
                row["llm_reasoning"] = r.llm_reasoning
            writer.writerow(row)
    logger.info(f"Results written to {output_path}")


# ── Main evaluation loop ────────────────────────────────────────────────────

def run_evaluation(csv_path: str, output_path: str, verbose: bool = False, scorer: str = "text"):
    # Validate env vars
    missing = [v for v in REQUIRED_ENV_VARS if not environ.get(v)]
    if missing:
        logger.error(f"Missing environment variables: {', '.join(missing)}")
        logger.error("Create a .env file based on .env.template")
        sys.exit(1)

    # Validate LLM env vars if needed
    use_llm = scorer in ("llm", "both")
    aoai_client = None
    if use_llm:
        if not environ.get("AZURE_OPENAI_ENDPOINT") or not environ.get("AZURE_OPENAI_API_KEY"):
            logger.error("LLM scorer requires AZURE_OPENAI_ENDPOINT and AZURE_OPENAI_API_KEY in .env")
            sys.exit(1)
        aoai_client = _get_aoai_client()
        logger.info(f"LLM judge: Azure OpenAI deployment '{environ.get('AZURE_OPENAI_DEPLOYMENT', 'gpt-4o')}'")

    use_text = scorer in ("text", "both")

    # Load test cases
    cases = load_test_cases(csv_path)
    if not cases:
        logger.error("No test cases to run")
        sys.exit(1)

    # Authenticate
    logger.info("Authenticating with Azure AD...")
    access_token = acquire_token()

    # Create client and start conversation
    logger.info("Connecting to Copilot Studio agent...")
    client = create_copilot_client(access_token)

    loop = asyncio.new_event_loop()
    asyncio.set_event_loop(loop)
    conversation_id = loop.run_until_complete(start_conversation(client))
    logger.info(f"Conversation started: {conversation_id}")

    results: list[TestResult] = []
    summary = EvalSummary(total=len(cases))
    llm_pass_count = 0
    llm_fail_count = 0
    llm_scores = []

    scorer_label = {"text": "Text Similarity", "llm": "LLM Judge", "both": "Text + LLM"}[scorer]

    print(f"\n{'='*70}")
    print(f"  Copilot Studio Agent Evaluation")
    print(f"  Scorer: {scorer_label}")
    print(f"  Test cases: {len(cases)}")
    print(f"  Started: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}")
    print(f"{'='*70}\n")

    for i, case in enumerate(cases, start=1):
        print(f"[{i}/{len(cases)}] Query: {case.query[:80]}{'...' if len(case.query) > 80 else ''}")

        start_time = time.time()
        error = ""
        actual_answer = ""

        try:
            actual_answer = loop.run_until_complete(
                ask_question(client, case.query, conversation_id)
            )
        except Exception as e:
            error = str(e)
            logger.error(f"  Error querying agent: {e}")
            summary.errors += 1

        latency = time.time() - start_time

        # ── Text-similarity scoring ──
        text_passed = ""
        confidence = 0.0
        if error:
            text_passed = "ERROR"
        elif not case.expected_answer:
            text_passed = "N/A"
        elif use_text:
            confidence = compute_confidence(case.expected_answer, actual_answer)
            text_passed = "PASS" if confidence >= PASS_THRESHOLD else "FAIL"

        # ── LLM scoring ──
        llm_score_val = -1.0
        llm_reasoning = ""
        llm_passed = ""
        if use_llm and case.expected_answer and not error:
            try:
                llm_score_val, llm_reasoning = llm_judge_score(
                    case.query, case.expected_answer, actual_answer, aoai_client
                )
                llm_passed = "PASS" if llm_score_val >= PASS_THRESHOLD else "FAIL"
                llm_scores.append(llm_score_val)
                if llm_passed == "PASS":
                    llm_pass_count += 1
                else:
                    llm_fail_count += 1
            except Exception as e:
                logger.error(f"  LLM judge error: {e}")
                llm_passed = "ERROR"
                llm_reasoning = str(e)
        elif not case.expected_answer:
            llm_passed = "N/A" if use_llm else ""
        elif error:
            llm_passed = "ERROR" if use_llm else ""

        # ── Determine primary pass/fail for summary ──
        if error:
            passed = "ERROR"
        elif not case.expected_answer:
            passed = "N/A"
            summary.skipped += 1
        else:
            # Use whichever scorer is active; if both, use LLM as primary
            if scorer == "llm":
                passed = llm_passed
            elif scorer == "both":
                passed = llm_passed  # LLM is primary when both are used
            else:
                passed = text_passed

            if passed == "PASS":
                summary.passed += 1
            elif passed == "FAIL":
                summary.failed += 1
            summary.evaluated += 1

        result = TestResult(
            row_number=case.row_number,
            query=case.query,
            expected_answer=case.expected_answer,
            actual_answer=actual_answer,
            passed=passed if use_text else (text_passed or passed),
            confidence_score=confidence,
            latency_seconds=latency,
            error=error,
            llm_score=llm_score_val,
            llm_reasoning=llm_reasoning,
            llm_passed=llm_passed,
        )
        results.append(result)

        # Print result line
        status_map = {"PASS": "+", "FAIL": "x", "N/A": "-", "ERROR": "!", "": "-"}
        print(f"  ", end="")
        if use_text and case.expected_answer and not error:
            icon = status_map[text_passed]
            print(f"[{icon}] Text: {text_passed} ({confidence:.2%})", end="")
            if use_llm:
                print("  |  ", end="")
        if use_llm and case.expected_answer and not error:
            icon = status_map.get(llm_passed, "-")
            score_str = f"{llm_score_val:.2%}" if llm_score_val >= 0 else "err"
            print(f"[{icon}] LLM: {llm_passed} ({score_str})", end="")
        if not case.expected_answer:
            print(f"[-] N/A (no expected answer)", end="")
        if error:
            print(f"[!] ERROR", end="")
        print(f"  [{latency:.1f}s]")

        if verbose and actual_answer:
            print(f"  Answer: {actual_answer[:120]}{'...' if len(actual_answer) > 120 else ''}")
        if verbose and llm_reasoning:
            print(f"  LLM reasoning: {llm_reasoning[:120]}{'...' if len(llm_reasoning) > 120 else ''}")
        if verbose and error:
            print(f"  Error: {error}")
        print()

    loop.close()

    # Compute summary averages
    text_scored = [r for r in results if r.passed in ("PASS", "FAIL")]
    if text_scored and use_text:
        summary.avg_confidence = sum(r.confidence_score for r in text_scored) / len(text_scored)
    summary.avg_latency = sum(r.latency_seconds for r in results) / len(results) if results else 0

    # Write output
    write_results(results, output_path, scorer)

    # Print summary
    print(f"\n{'='*70}")
    print(f"  EVALUATION SUMMARY  (scorer: {scorer_label})")
    print(f"{'='*70}")
    print(f"  Total test cases:    {summary.total}")
    print(f"  Evaluated (w/ exp.): {summary.evaluated}")
    print(f"  Skipped (no exp.):   {summary.skipped}")
    print(f"  Errors:              {summary.errors}")

    if use_text:
        text_pass = sum(1 for r in results if r.passed == "PASS" or (scorer == "llm" and r.confidence_score >= PASS_THRESHOLD))
        print(f"\n  --- Text Similarity ---")
        text_scored_results = [r for r in results if r.confidence_score > 0 or (r.expected_answer and not r.error)]
        if text_scored_results:
            tp = sum(1 for r in results if r.expected_answer and not r.error and r.confidence_score >= PASS_THRESHOLD)
            tf = sum(1 for r in results if r.expected_answer and not r.error and r.confidence_score < PASS_THRESHOLD)
            avg_c = summary.avg_confidence
            print(f"  Passed:              {tp}")
            print(f"  Failed:              {tf}")
            print(f"  Avg confidence:      {avg_c:.2%}")

    if use_llm:
        print(f"\n  --- LLM Judge ---")
        print(f"  Passed:              {llm_pass_count}")
        print(f"  Failed:              {llm_fail_count}")
        if llm_scores:
            print(f"  Avg LLM score:       {sum(llm_scores)/len(llm_scores):.2%}")

    print(f"\n  Avg latency:         {summary.avg_latency:.1f}s")
    print(f"  Results saved to: {output_path}")
    print(f"{'='*70}\n")


# ── CLI entry point ──────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(
        description="Evaluate a Copilot Studio agent against a CSV test suite."
    )
    parser.add_argument(
        "input_csv",
        help="Path to CSV file with 'query' and 'expected_answer' columns",
    )
    parser.add_argument(
        "--output", "-o",
        default=None,
        help="Output CSV path (default: <input>_results_<timestamp>.csv)",
    )
    parser.add_argument(
        "--verbose", "-v",
        action="store_true",
        help="Show actual answers and errors in console output",
    )
    parser.add_argument(
        "--threshold", "-t",
        type=float,
        default=0.6,
        help="Confidence threshold for PASS (default: 0.6)",
    )
    parser.add_argument(
        "--scorer", "-s",
        choices=["text", "llm", "both"],
        default="text",
        help="Scoring method: 'text' (text similarity), 'llm' (Azure OpenAI judge), 'both' (default: text)",
    )

    args = parser.parse_args()

    global PASS_THRESHOLD
    PASS_THRESHOLD = args.threshold

    if args.output:
        output_path = args.output
    else:
        stem = Path(args.input_csv).stem
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_path = f"{stem}_results_{timestamp}.csv"

    run_evaluation(args.input_csv, output_path, verbose=args.verbose, scorer=args.scorer)


if __name__ == "__main__":
    main()
