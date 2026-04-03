# Copilot Studio Agent Evaluation Utility

A command-line tool for automated testing and evaluation of Microsoft Copilot Studio agents. Feed it a CSV of test queries and expected answers, and it will call your agent, compare responses, and produce a scored results report.

## Features

- **Automated agent testing** — sends queries from a CSV file to your Copilot Studio agent via the Microsoft 365 Agent SDK
- **Delegated authentication** — opens a browser for Microsoft login (OAuth 2.0 authorization code flow)
- **Two scoring methods:**
  - **Text Similarity** — fast, local scoring using sequence matching and keyword overlap
  - **LLM Judge** — semantic scoring via Azure OpenAI that understands meaning, not just word matching
- **Detailed results CSV** — per-query pass/fail, confidence scores, latency, LLM reasoning
- **Console summary** — pass rate, average confidence, average latency

## Prerequisites

- Python 3.10+
- An Azure AD app registration with:
  - `CopilotStudio.Copilots.Invoke` delegated permission
  - A client secret
  - `http://localhost:8400/auth/callback` added as a **Web** redirect URI
- A published Copilot Studio agent
- (Optional) An Azure OpenAI resource for LLM-based scoring

## Setup

1. **Install dependencies:**

   ```bash
   pip install -r requirements.txt
   ```

2. **Configure environment variables:**

   Copy the template and fill in your values:

   ```bash
   cp .env.template .env
   ```

   Required variables:

   | Variable | Description |
   |----------|-------------|
   | `COPILOTSTUDIOAGENT__AGENTAPPID` | Azure AD application (client) ID |
   | `COPILOTSTUDIOAGENT__CLIENTSECRET` | Azure AD client secret |
   | `COPILOTSTUDIOAGENT__TENANTID` | Azure AD tenant ID |
   | `COPILOTSTUDIOAGENT__ENVIRONMENTID` | Power Platform environment ID (from Copilot Studio settings) |
   | `COPILOTSTUDIOAGENT__SCHEMANAME` | Agent schema name (from Copilot Studio metadata) |

   Additional variables for LLM scoring (`--scorer llm` or `--scorer both`):

   | Variable | Description |
   |----------|-------------|
   | `AZURE_OPENAI_ENDPOINT` | Azure OpenAI resource endpoint URL |
   | `AZURE_OPENAI_API_KEY` | Azure OpenAI API key |
   | `AZURE_OPENAI_DEPLOYMENT` | Deployment name (default: `gpt-4o`) |
   | `AZURE_OPENAI_API_VERSION` | API version (default: `2024-12-01-preview`) |

## CSV Format

Create a CSV file with two columns:

```csv
query,expected_answer
What are your business hours?,Monday to Friday 9am to 5pm
Tell me about returns,
I need help with a contract,
```

- **`query`** (required) — the question to send to the agent
- **`expected_answer`** (optional) — leave blank to capture the response without scoring it

## Usage

```bash
# Basic run with text similarity scoring (default)
python evaluate.py tests.csv

# Use LLM judge for semantic scoring
python evaluate.py tests.csv --scorer llm

# Run both scorers side by side
python evaluate.py tests.csv --scorer both

# Verbose output with custom threshold and output file
python evaluate.py tests.csv --scorer both --verbose --threshold 0.7 --output my_results.csv
```

### CLI Options

| Flag | Short | Description |
|------|-------|-------------|
| `input_csv` | | Path to the CSV test file |
| `--output` | `-o` | Output CSV path (default: `<input>_results_<timestamp>.csv`) |
| `--scorer` | `-s` | Scoring method: `text`, `llm`, or `both` (default: `text`) |
| `--threshold` | `-t` | Confidence threshold for PASS/FAIL (default: `0.6`) |
| `--verbose` | `-v` | Show actual answers and LLM reasoning in console |

## Scoring Methods

### Text Similarity (`--scorer text`)

A fast, local scorer that doesn't require any external API. It uses:

| Check | Score | Description |
|-------|-------|-------------|
| Exact match | 1.0 | Normalized text is identical |
| Containment | 0.95 | Expected answer appears verbatim in actual response |
| Fuzzy blend | 0.0 - 0.94 | 60% SequenceMatcher ratio + 40% keyword overlap |

**Best for:** exact or near-exact answer validation, keyword checking.

### LLM Judge (`--scorer llm`)

Sends each (query, expected answer, actual answer) to Azure OpenAI and asks it to score semantic correctness from 0.0 to 1.0, along with a brief reasoning.

**Best for:** answers where phrasing varies but meaning should match. For example, "Visit the password reset page" and "Go to Settings > Security > Reset Password" would score high with LLM but lower with text similarity.

### Both (`--scorer both`)

Runs both scorers and includes all columns in the output. The LLM score is used as the primary pass/fail when both are active.

## Output

The results CSV includes:

| Column | When included |
|--------|---------------|
| `row_number` | Always |
| `query` | Always |
| `expected_answer` | Always |
| `actual_answer` | Always |
| `text_passed` | `text` or `both` |
| `text_confidence` | `text` or `both` |
| `llm_passed` | `llm` or `both` |
| `llm_score` | `llm` or `both` |
| `llm_reasoning` | `llm` or `both` |
| `latency_seconds` | Always |
| `error` | Always |

A console summary is also printed at the end of each run with aggregate statistics.

## Example

```
======================================================================
  Copilot Studio Agent Evaluation
  Scorer: Text + LLM
  Test cases: 3
  Started: 2026-04-03 16:50:00
======================================================================

[1/3] Query: What are your business hours?
  [+] Text: PASS (82.50%)  |  [+] LLM: PASS (90.00%)  [2.3s]

[2/3] Query: Tell me about returns
  [-] N/A (no expected answer)  [1.8s]

[3/3] Query: I need help with a contract
  [-] N/A (no expected answer)  [2.1s]

======================================================================
  EVALUATION SUMMARY  (scorer: Text + LLM)
======================================================================
  Total test cases:    3
  Evaluated (w/ exp.): 1
  Skipped (no exp.):   2
  Errors:              0

  --- Text Similarity ---
  Passed:              1
  Failed:              0
  Avg confidence:      82.50%

  --- LLM Judge ---
  Passed:              1
  Failed:              0
  Avg LLM score:       90.00%

  Avg latency:         2.1s
  Results saved to: tests_results_20260403_165012.csv
======================================================================
```

## Project Structure

```
CPSTesting/
├── evaluate.py          # Main evaluation script
├── requirements.txt     # Python dependencies
├── .env                 # Environment variables (do not commit)
├── sample_tests.csv     # Example test CSV
└── README.md
```
