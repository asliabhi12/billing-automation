import pandas as pd
import json
import unicodedata

def normalize_text(x):
    if pd.isna(x):
        return ""
    
    x = str(x)

    # Normalize unicode characters
    x = unicodedata.normalize("NFKC", x)

    # Replace strange hyphens with normal hyphen
    x = (
        x.replace("\u2011", "-")   # non-breaking hyphen
         .replace("\u2013", "-")   # en dash
         .replace("\u2014", "-")   # em dash
         .replace("\u2212", "-")   # minus sign
    )

    # Replace non-breaking spaces
    x = x.replace("\u00a0", " ")

    return x.strip()


df = pd.read_excel("customer_mapping.xlsx", engine="openpyxl")

df["SubscriptionId"] = df["SubscriptionId"].apply(normalize_text)
df["Customer Name"] = df["Customer Name"].apply(normalize_text)

# remove duplicates safely
df = df.drop_duplicates(subset="SubscriptionId", keep="first")

mapping = dict(zip(df["SubscriptionId"], df["Customer Name"]))

with open("customer_map.json", "w", encoding="utf-8") as f:
    json.dump(mapping, f, indent=2, ensure_ascii=False)

print("Mapping file created successfully.")