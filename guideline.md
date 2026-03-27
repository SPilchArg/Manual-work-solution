# Guideline - Indeed Proto QA Reviewer

## Objective

Run a fast QA prototype for DOCX articles against reference materials and internal spot checks using a 2-agent workflow (Gemini then Claude).

## Standard Workflow

1. Launch `launch_gui.bat`.
2. Select folders:
   - Articles
   - References
   - Internal spot checks
3. Select `.env` containing Google + Anthropic API keys.
4. Select output folder.
5. Run `Run 2-Agent QA Workflow`.

## Expected Outputs

- Annotated article DOCX files in `annotated_articles/`
- `issues_summary.docx`
- `qa_readiness_report.xlsx`
- `qa_rules.json`
- `issues.json`
- Extracted JSON in `extracted/`

## Quality Checks

- Confirm each input folder contains `.docx` files.
- Confirm `.env` has both keys.
- Verify score and issue totals in the Excel report.
- Spot-check one annotated DOCX for highlighted segments and QA comment block.

## Troubleshooting

- If launcher fails on dependency install, ensure Python/pip are in PATH.
- If API calls fail, app switches to fallback mode; verify internet access and key validity.
- If no article files are found, ensure files are `.docx` and not temporary lock files (`~$...`).

## Prototype Limits

- Highlight/comment behavior is simplified for speed.
- LLM response quality depends on prompt size and model/API availability.
- Cost sheet uses requested prototype formula.
