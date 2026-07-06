import streamlit as st
import requests

# Secrets uit secrets.toml
client_id = st.secrets["client_id"]
client_secret = st.secrets["client_secret"]
tenant_id = st.secrets["tenant_id"]
excel_filename = st.secrets["excel_filename"]
table_name = st.secrets["table_name"]

def get_token():
    url = f"https://login.microsoftonline.com/{tenant_id}/oauth2/v2.0/token"
    data = {
        "client_id": client_id,
        "client_secret": client_secret,
        "scope": "https://graph.microsoft.com/.default",
        "grant_type": "client_credentials"
    }
    r = requests.post(url, data=data)
    r.raise_for_status()
    return r.json()["access_token"]

def find_excel_file(token):
    url = f"https://graph.microsoft.com/v1.0/me/drive/root/search(q='{excel_filename}')"
    r = requests.get(url, headers={"Authorization": f"Bearer {token}"})
    r.raise_for_status()
    items = r.json().get("value", [])
    if not items:
        raise RuntimeError(f"Bestand {excel_filename} niet gevonden in OneDrive")
    return items[0]["parentReference"]["driveId"], items[0]["id"]

st.title("HR KPI – maandelijkse rapportage")

# Vaste vragen
questions = [
    "Number of open vacancies",
    "Number of vacancies filled",
    "Number of employees started this month",
    "Are there any recruitment challenges or hiring issues",
    "Total actual working hours",
    "Total medical leave",
    "Any employees that may present a retention risk",
    "Any concerns related to engagement",
    "Any concerns related to turnover",
    "Any concerns related to employee satisfaction",
    "Relevant updates or concerns regarding team leads",
    "Relevant updates or concerns regarding other HR-related matters"
]

# Algemene velden
country = st.text_input("Reporting on country")
reporting_period = st.text_input("Reporting on period")

# Dynamische invoervelden voor alle vaste vragen
answers = {}
for q in questions:
    answers[q] = st.text_input(q)

if st.button("Opslaan"):
    try:
        token = get_token()
        drive_id, item_id = find_excel_file(token)

        # Voor elke vraag één rij toevoegen
        for q in questions:
            row = {
                "values": [[
                    country,
                    reporting_period,
                    q,
                    answers[q]
                ]]
            }

            url = (
                f"https://graph.microsoft.com/v1.0/"
                f"drives/{drive_id}/items/{item_id}/workbook/tables/{table_name}/rows/add"
            )

            r = requests.post(
                url,
                headers={"Authorization": f"Bearer {token}"},
                json=row
            )
            r.raise_for_status()

        st.success("Alle antwoorden zijn opgeslagen!")
    except Exception as e:
        st.error(f"Fout: {e}")
