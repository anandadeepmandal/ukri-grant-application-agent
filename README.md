# UKRI Grant Application Agent

Developed by a.mandal@bham.ac.uk

This web app lets a user paste a public profile URL or profile text, searches UKRI for relevant funding opportunities, proposes topics matched to the user's expertise, and generates a downloadable Word application draft.

## Workflow

1. Paste profile URL or profile text.
2. Search UKRI open and upcoming funding opportunities.
3. Select an interesting grant.
4. Generate project topics matched to the applicant's profile.
5. Select a topic.
6. Enter applicant details.
7. Download a Word application draft.

## Render settings

Runtime: Python

Build Command:

```bash
pip install -r requirements.txt
```

Start Command:

```bash
uvicorn app:app --host 0.0.0.0 --port $PORT
```

Environment variable if needed:

```text
PYTHON_VERSION=3.11.11
```

Always check the official UKRI page, eligibility, attachments, costing rules and institutional approval requirements before final submission.
