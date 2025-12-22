import requests
from config.itsperfect import BASE_URL, USERNAME, PASSWORD

def get_itsperfect_token():
    r = requests.post(
        f"{BASE_URL}/authentication",
        json={"username": USERNAME, "password": PASSWORD}
    )
    r.raise_for_status()
    return r.json()["token"]