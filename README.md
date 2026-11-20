# Share Class Data Manager

A deployment-ready Streamlit demo for browsing and managing the genericized share-class master data
in `data/Genericized_Sample_Data.xlsx`.

## Demo features

- Dashboard with record and data-quality metrics
- HTML-prototype-inspired navy and maroon theme with top-level horizontal navigation
- Overview table with operational filters, field-group selection, and coverage drilldowns
- Search and filters by lifecycle, currency, and share-class type
- Grouped add/edit form covering all 37 taxonomy fields
- Form entries remain available after submission, including when validation fails
- Reviewer queue for record, deletion, and workbook-import proposals
- Automatic ISIN generation: `GF555555555<next integer>`
- Duplicate, required-field, date, numeric, and lock-rule validation
- Record proposals are blocked by duplicate identifiers belonging to the proposed record,
  not by unrelated duplicates elsewhere in the dataset
- Corrected taxonomy and working-data downloads
- Session audit history
- Demo roles for Viewer, Editor, Reviewer, and Admin

## Run locally

```powershell
py -m venv .venv
.venv\Scripts\Activate.ps1
pip install -r requirements.txt
streamlit run app.py
```

## Deploy for a free demo

1. Create a GitHub repository and push this project folder.
2. Go to Streamlit Community Cloud and choose **Create app**.
3. Select the repository and branch.
4. Set the main file path to `app.py`.
5. Deploy. Streamlit Cloud builds the environment from `requirements.txt`.

The application deliberately has no compiled or OS-specific dependencies. It includes only
genericized sample data.

## Important demo limitations

Changes are held in `st.session_state` for the current browser session and do not survive an app
restart. Download the updated workbook to preserve the demonstration result.

The role selector demonstrates authorization behavior; it is **not authentication**. Do not use this
deployment with confidential or regulated data.

## Production direction

Before handling real data:

- Authenticate users through the organization's approved Microsoft Entra ID.
- Require an explicit company-domain identity and Entra group membership.
- Suggested groups: `ShareClass-Viewers`, `ShareClass-Editors`,
  `ShareClass-Reviewers`, and `ShareClass-Admins`.
- Enforce role permissions in every server-side mutation.
- Store approved records and audit events in an authenticated managed database.
- Implement atomic ISIN allocation in the database to handle concurrent users.
- Add retention, logging, monitoring, backups, and an approved private deployment boundary.

## Project structure

```text
app.py                     Streamlit application and workflows
share_class_manager/       Schema, XLSX I/O, validation, and ISIN allocation
data/                      Genericized source and corrected taxonomy
tests/                     Core behavior tests
.streamlit/config.toml     UI and server configuration
requirements.txt          Deployment dependencies
```

## Validate

```powershell
python -m unittest discover -s tests -v
```

