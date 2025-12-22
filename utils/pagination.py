import time
import requests
from utils.auth import get_itsperfect_token
from tqdm import tqdm

def fetch_paginated(url, headers, limit=250):
    all_data = []

    # First, get total pages (optional)
    r = requests.get(f"{url}&limit={limit}&page=1", headers=headers)
    r.raise_for_status()
    total_pages = int(r.headers.get("X-Pagination-Page-Count", 1))
    
    for page in tqdm(range(1, total_pages + 1), desc="Fetching pages"):
        while True:
            r = requests.get(f"{url}&limit={limit}&page={page}", headers=headers)
            print(f"Fetched {len(all_data)} orders so far...")
            if r.status_code == 429:
                # rate limit handling
                time.sleep(4)
                continue
            elif r.status_code == 401:
                headers["Authorization"] = f"Bearer {get_itsperfect_token()}"
                continue
            r.raise_for_status()
            break

        data = r.json()
        all_data.extend(data)

    return all_data