import pandas as pd

# Load Excel file
file_path = "DLP CERTIFICATE.xlsx"
df = pd.read_excel(file_path)

# CHANGE THIS to the exact column name that contains names
# set to a lowercase-friendly default (will be normalized)
NAME_COLUMN = "name"   # example: "STUDENT NAME"

# Normalize column names to help with inconsistent headers
df.rename(columns=lambda c: str(c).strip().lower(), inplace=True)

# Resolve name column (allow common variants)
NAME_COLUMN = NAME_COLUMN.strip().lower()
if NAME_COLUMN not in df.columns:
    common = ["name", "full name", "student name", "fullname", "first name"]
    found = next((c for c in common if c in df.columns), None)
    if not found:
        found = next((c for c in df.columns if "name" in c), None)
    if found:
        name_col = found
    else:
        raise KeyError(f"No name column found. Available columns: {list(df.columns)}")
else:
    name_col = NAME_COLUMN


def proper_name(name):
    if pd.isna(name):
        return name

    words = str(name).split()
    formatted_words = []

    for word in words:
        # If word is all caps and length <= 4, keep it uppercase (PV, AI, CSE)
        if word.isupper() and len(word) <= 4:
            formatted_words.append(word)
        else:
            formatted_words.append(word.capitalize())

    return " ".join(formatted_words)

# Apply formatting
df[name_col] = df[name_col].apply(proper_name)

# Save new file
output_file = "DLP_CERTIFICATE_Proper_Names.xlsx"
df.to_excel(output_file, index=False)

print("✅ Bulk name formatting completed!")
print(f"📄 Saved as: {output_file}")
