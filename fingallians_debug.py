"""Quick debug script — run this and paste the output back to Claude."""
import requests

API_URL = "https://dublingaa.sportlomo.com/wp-admin/admin-ajax.php"
headers = {
    "User-Agent": "Mozilla/5.0 (Macintosh; Intel Mac OS X 10_15_7) AppleWebKit/537.36",
    "Referer": "https://dublingaa.sportlomo.com/clubprofile/3573",
}

params = {
    "action": "club_page_results",
    "org": "3",
    "club_id": "3573",
    "sport": "",
    "competition_id": "",
    "team_id": "",
    "user_id": "3",
    "fdate": "2026-01-01",
    "tdate": "2026-12-31",
    "age": "",
    "nosuper": "",
    "displayResults": "",
}

resp = requests.get(API_URL, params=params, headers=headers, timeout=30)
print(f"Status: {resp.status_code}")
print(f"Content-Type: {resp.headers.get('Content-Type')}")
# Print first 4000 chars of raw HTML
print("\n--- RAW HTML (first 4000 chars) ---")
print(resp.text[:4000])
