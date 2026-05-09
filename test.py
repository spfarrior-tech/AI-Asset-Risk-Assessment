import pandas as pd
df = pd.read_excel('Assets.xlsx')
condition_scores = {
    "Poor": 5,
    "Fair": 3,
    "Good": 1
}
failure_scores = {
    "High": 5,
    "Medium": 3,
    "Low": 1
}
df['Condition_Score'] = df['Condition'].map(condition_scores)
df['Failure_Score'] = df['Failure_Risk'].map(failure_scores)

df["Risk_Score"] = (
    df["Age"] * 0.3 +
    df["Condition_Score"] * 2 +
    df["Failure_Score"] * 2 +
    df["Replacement_Cost"] / 100000
)
df = df.sort_values(by="Risk_Score", ascending=False)

print(df[[
    "Asset_ID",
    "Asset_Name",
    "Condition",
    "Age",
    "Replacement_Cost",
    "Failure_Risk",
    "Risk_Score"
]].head(10))

df.to_excel('Asset_Risk_Assessment.xlsx', index=False)
print("Report created: Asset_Risk_Assessment.xlsx")

from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib import colors
from datetime import datetime

pdf_file = "Asset_Risk_Assessment.pdf"

doc = SimpleDocTemplate(pdf_file, pagesize=letter)
styles = getSampleStyleSheet()
story = []

title = Paragraph("Asset Risk Assessment Report", styles['Title'])
story.append(title)
story.append(Spacer(1, 12))

date_text = Paragraph(f"Date Generated: {datetime.today().strftime('%B')} {datetime.today().day}, {datetime.today().year}", styles['Normal'])
story.append(date_text)
story.append(Spacer(1, 18))

summary = Paragraph(
    "Executive Summary: This report identifies the top high risk assets based on age, condition, failure risk, and replacement cost. Assets with higher risk scores should be prioritized for inspection, maintenance planning, and capital investment review.",
    styles['BodyText']
)
story.append(summary)
story.append(Spacer(1, 18))

top_10 = df[[
    "Asset_ID",
    "Asset_Name",
    "Condition",
    "Age",
    "Replacement_Cost",
    "Failure_Risk",
    "Risk_Score"
]].head(10).copy()

top_10['Replacement_Cost'] = top_10['Replacement_Cost'].apply(lambda x: f"${x:,.0f}")
top_10['Risk_Score'] = top_10['Risk_Score'].round(0).astype(int)
table_data = [top_10.columns.tolist()] + top_10.values.tolist()
table = Table(table_data, repeatRows=1)

table.setStyle(TableStyle([
    ('BACKGROUND', (0, 0), (-1, 0), colors.lightgrey),
    ('TEXTCOLOR', (0, 0), (-1, 0), colors.whitesmoke),
    ('GRID', (0, 0), (-1, -1), 0.5, colors.grey),
    ('FONTNAME', (0, 0), (-1, 0), 'Helvetica-Bold'),
    ('FONTSIZE', (0, 0), (-1, -1), 8),
    ('ALIGN', (0, 1), (-1, -1), "CENTER"),
    ('VALIGN', (0, 0), (-1, -1), "MIDDLE"),
]))

story.append(Paragraph("Top 10 High Risk Assets:", styles['Heading2']))
story.append(Spacer(1, 8))
story.append(table)
story.append(Spacer(1, 18))

recommendations = Paragraph(
    "Recommendations: Leadership should review the highest-ranked assets first and determine whether "
    "they require immediate attention, maintenance, or replacement forescasting, or capital "
    "planning consieration.",
    styles['BodyText']
)
story.append(recommendations)

doc.build(story)

print("PDF created: Asset_Risk_Assessment_Report.pdf")