import os, re, time
from ddgs import DDGS
import requests
from openpyxl import Workbook

QUERIES = [
    "engineering competition South Africa",
    "innovation exhibition South Africa",
    "engineering design challenge South Africa",
    "university engineering competition South Africa",
    "tech expo South Africa",
    "maker faire South Africa",
    "stem competition South Africa",
    "robotics competition South Africa",
    "civil engineering competition South Africa",
    "mechanical engineering competition South Africa",
    "electrical engineering competition South Africa",
    "engineering hackathon South Africa",
]

OUTPUT = "data/competitions_with_emails.xlsx"

EMAIL_REGEX = re.compile(r'[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}')

def extract_emails(url):
    """Visit a URL and extract all email addresses from the HTML."""
    emails = set()
    try:
        headers = {'User-Agent': 'Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36'}
        response = requests.get(url, timeout=10, headers=headers)
        if response.status_code == 200:
            emails.update(EMAIL_REGEX.findall(response.text))
    except Exception:
        pass
    return ', '.join(emails) if emails else ''

def main():
    os.makedirs("data", exist_ok=True)
    wb = Workbook()
    ws = wb.active
    ws.title = "Competitions"
    ws.append(["Title", "Link", "Snippet", "Source Query", "Emails Found"])

    total = 0
    with DDGS() as ddgs:
        for q in QUERIES:
            print(f"🔍 Searching: {q}")
            try:
                for result in ddgs.text(q, max_results=10):
                    title = result.get("title", "")
                    link = result.get("href", "")
                    snippet = result.get("body", "")
                    print(f"   📄 {title[:60]}...")
                    emails = extract_emails(link)
                    ws.append([title, link, snippet, q, emails])
                    total += 1
                    time.sleep(0.5)   # be gentle
            except Exception as e:
                print(f"   ⚠️ {e}")

    wb.save(OUTPUT)
    print(f"\n✅ Saved {total} results to {OUTPUT}")

if __name__ == "__main__":
    main()
