# feasibility-report-generator

Tkinter desktop application for generating and refining feasibility report content with OpenAI, then filling a Word template.

## Requirements

- Python 3.10+
- `OPENAI_API_KEY` environment variable
- `Feasibility_Report_Word_Template.docx` in the repository root

Install dependencies:

```bash
pip install -r requirements.txt
```

## Run

```bash
python generate_report.py
```

## Workflow

1. Enter **Project Title**, **Author**, and **Use Case Description**.
2. Click **Generate Initial Report Content**.
3. For each section, edit directly and/or click **Improve** to request AI refinement.
4. Click **Approve** for each section once satisfied.
5. After all sections are approved, **Generate Report** is enabled to save the final `.docx`.

## Report sections

The app includes these sections (including new required sections):

- Overview
- Business Case
- Problem Definition
- Value Proposition
- Challenges
- Success Criteria
- Data Availability
- Input corpus
- Data quality
- Technical Feasibility
- Analytical complexity
- Infrastructure
- Integration
- Risk Assessment
- Mitigation strategies
- Recommendations
- Recommendation to Proceed
- Rationale
