import pandas as pd
import re

with open("messy-updates.txt", "r", encoding="utf-8") as file:
    lines = file.readlines()

data = []
current_date = None

for line in lines:
    line = line.strip()

    date_match = re.search(r"\d{2}-\d{2}-\d{4}", line)
    if date_match:
        current_date = date_match.group()
        continue

    if line and current_date:
        data.append({
            "Date": current_date,
            "Action": line,
            "Status": "Pending"
        })

df = pd.DataFrame(data)
df.to_excel("structured_actions.xlsx", index=False)

print("Excel file created successfully")
