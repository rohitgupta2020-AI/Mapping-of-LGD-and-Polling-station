# Polling Station LGD Manual Mapper

Streamlit app for manually mapping unmapped polling-station rows to LGD village or ward records.

## Features

- Login-protected access for small teams.
- Separate filters for polling constituency, LGD constituency, and LGD district.
- Select one or more polling rows and one LGD village/ward row.
- Writes the LGD village name, code, block, district, and entity type into the polling workbook.
- Downloads or saves a new timestamped Excel file without overwriting the source workbook.

## Local Setup

```powershell
pip install -r requirements.txt
Copy-Item .streamlit\secrets.example.toml .streamlit\secrets.toml
python scripts\hash_password.py "your-password"
streamlit run app.py
```

Put the generated hashes into `.streamlit/secrets.toml`.

## Login Users

Configure up to 10 users in Streamlit secrets:

```toml
[auth.users.user1]
password_hash = "sha256-password-hash"
```

Add more users as `[auth.users.user2]`, `[auth.users.user3]`, and so on.

## Data Files

For public GitHub deployment, do not commit real Excel files. They are ignored by `.gitignore`.

Users can upload these files after login:

- `Polling_Station_LGD mapping.xlsx`
- `LGD with constituency.xlsx`

If you deploy privately and want bundled files, place both workbooks beside `app.py` on the server.

## Deploy To Streamlit Community Cloud

1. Push this repository to GitHub.
2. Go to Streamlit Community Cloud and create a new app from the repository.
3. Set `app.py` as the main file.
4. Paste your `[auth.users...]` values into the app secrets.
5. Deploy.

The public app URL can be shared with your team. Each user signs in with their assigned login ID and password.
