import pandas as pd
import re

def normalize(text):
    return re.sub(r'\s+', ' ', str(text)).strip().lower()

def norma(text):
    return re.sub(r'\s+', ' ', str(text)).strip()

# Load data1 (main table)
data1 = pd.read_excel(
    r"C:\Users\PC\Downloads\Telegram Desktop\Delib_ING2-S3_24-25.xlsx",
    skiprows=4
)
normalized_data1 = data1.fillna('').applymap(normalize)

# Load data2 (name list)
data2 = pd.read_excel(
    r"C:\Users\PC\Downloads\rattrappage.xlsx",
    skiprows=4,
    header=None
)

# Combine nom + prénom (adjust column indices if needed)
data2['FullName'] = data2[1].astype(str) + ' ' + data2[2].astype(str)
data2['FullName'] = data2['FullName'].apply(normalize)

class StudentRattrapage:
    def __init__(self, N, name, UNID):
        self.N = N
        self.name = name
        self.UNID = UNID
        self.subjects = {}         # subject name -> grade < 10
        self.rattrapage = []       # list of subject names with grade < 10

    def add_subject(self, subject, grade):
        self.subjects[subject] = grade
        self.rattrapage.append(subject)

    def __str__(self):
        out = f"👤 {self.name.upper()} (N°: {self.N}, ID: {self.UNID})"
        if self.subjects:
            out += f"\n 🔻 Subjects {len(self.subjects)}: {', '.join(self.rattrapage)}"
            for subject, grade in self.subjects.items():
                out += f"\n   - {subject}: {grade}"
        else:
            out += "\n ✅ No failed subjects."
        return out

arr = []

for name in data2['FullName']:
    match = data1[normalized_data1.apply(
        lambda row: row.astype(str).str.contains(name).any(),
        axis=1
    )]

    if not match.empty:
        for _, row in match.iterrows():
            student = StudentRattrapage(
                N=row.iloc[0],
                name=norma(row.iloc[1]),
                UNID=row.iloc[2]
            )

            for col_name, value in row.iloc[3:11].items():
                try:
                    num = float(value)
                    if num < 10:
                        student.add_subject(col_name, num)
                except (ValueError, TypeError):
                    continue

            arr.append(student)

# ---- Print them with custom format ----
for student in arr:
    print("\n" + str(student))

with open("rattrapage_output.txt", "w", encoding="utf-8") as f:
    for student in arr:
        f.write(str(student) + "\n\n")