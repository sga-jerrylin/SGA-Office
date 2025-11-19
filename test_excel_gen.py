import requests
import json

url = "http://127.0.0.1:5101/generate_excel"

payload = {
    "filename": "Markdown_Excel_Test",
    "content": """
# Sales Data

Here is the sales data for Q1 and Q2.

| Product | Q1 Sales | Q2 Sales | Total |
| :--- | :---: | :---: | ---: |
| Widget A | 100 | 150 | 250 |
| Widget B | 200 | 250 | 450 |
| **Total** | **300** | **400** | **700** |

End of report.
"""
}

headers = {
  'Content-Type': 'application/json'
}

try:
    response = requests.post(url, headers=headers, data=json.dumps(payload))
    print(f"Status Code: {response.status_code}")
    if response.status_code == 200:
        print("Response:", json.dumps(response.json(), indent=2, ensure_ascii=False))
    else:
        print("Error:", response.text)
except Exception as e:
    print(f"Request failed: {e}")
