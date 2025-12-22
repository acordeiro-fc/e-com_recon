import requests

BASE_URL = st.secrets["ITSP_BASE_URL"]
USERNAME = st.secrets["ITSP_USERNAME"]
PASSWORD = st.secrets["ITSP_PASSWORD"]

def get_itsperfect_token():
    r = requests.post(
        f"{BASE_URL}/authentication",
        json={"username": USERNAME, "password": PASSWORD}
    )
    r.raise_for_status()

    return r.json()["token"]
