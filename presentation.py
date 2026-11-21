
import pandas as pd
import matplotlib.pyplot as plt
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor

# ==========================================
# 1. LOAD THE DATA
# ==========================================

df = pd.read_csv("nutrition_data.csv")

print("Dataset loaded successfully!")
print(df.head())

print("\nDataset shape:")
print(df.shape)

print("\nColumn names:")
print(df.columns.tolist())

# ==========================================
# 2. ANALYSE THE DATA
# ==========================================

# Average weight gain by county
county_analysis = df.groupby("county")["weight_gain_kg"].mean()

print("\nAVERAGE WEIGHT GAIN BY COUNTY")
print(county_analysis.round(2))


# Number of beneficiaries by county
county_counts = df["county"].value_counts()

print("\nBENEFICIARIES BY COUNTY")
print(county_counts)


# Average weight gain by age group
age_analysis = df.groupby("age_group")["weight_gain_kg"].mean()

print("\nAVERAGE WEIGHT GAIN BY AGE GROUP")
print(age_analysis.round(2))


# Relationship between attendance and weight gain
correlation = df["sessions_attended"].corr(
    df["weight_gain_kg"]
)

print("\nATTENDANCE AND WEIGHT GAIN CORRELATION")
print(round(correlation, 2))


# Overall programme statistics
print("\nOVERALL PROGRAMME STATISTICS")
print("Total beneficiaries:", len(df))
print("Average weight gain:", round(df["weight_gain_kg"].mean(), 2), "kg")
print("Average sessions attended:", round(df["sessions_attended"].mean(), 1))

# ==========================================
# 3. CREATE CHARTS FOR THE PRESENTATION
# ==========================================

import os

# Create a folder for our charts
os.makedirs("charts", exist_ok=True)


# CHART 1: Average weight gain by county

county_analysis.sort_values().plot(
    kind="bar",
    figsize=(8, 5),
    title="Average Weight Gain by County",
    color="steelblue"
)

plt.ylabel("Average Weight Gain (kg)")
plt.xlabel("County")
plt.xticks(rotation=0)
plt.tight_layout()
plt.savefig("charts/weight_gain_by_county.png")
plt.close()


# CHART 2: Average weight gain by age group

age_analysis.sort_values().plot(
    kind="bar",
    figsize=(8, 5),
    title="Average Weight Gain by Age Group",
    color="seagreen"
)

plt.ylabel("Average Weight Gain (kg)")
plt.xlabel("Age Group")
plt.xticks(rotation=0)
plt.tight_layout()
plt.savefig("charts/weight_gain_by_age.png")
plt.close()


# CHART 3: Attendance versus weight gain

plt.figure(figsize=(8, 5))

plt.scatter(
    df["sessions_attended"],
    df["weight_gain_kg"],
    alpha=0.7
)

plt.title("Programme Attendance vs Weight Gain")
plt.xlabel("Sessions Attended")
plt.ylabel("Weight Gain (kg)")

plt.text(
    0.05, 0.95,
    f"Correlation: {correlation:.2f}",
    transform=plt.gca().transAxes,
    verticalalignment="top"
)

plt.tight_layout()
plt.savefig("charts/attendance_vs_weight_gain.png")
plt.close()

print("\nAll three charts created successfully!")

# ==========================================
# 4. CREATE THE POWERPOINT PRESENTATION
# ==========================================

# Create a new PowerPoint presentation
prs = Presentation()

# Set widescreen slide dimensions (16:9)
prs.slide_width = Inches(13.333)
prs.slide_height = Inches(7.5)

# Corporate colours
DARK_BLUE = RGBColor(25, 50, 85)
WHITE = RGBColor(255, 255, 255)
GREY = RGBColor(90, 90, 90)


# Function to create a slide title
def add_title(slide, title):

    box = slide.shapes.add_textbox(
        Inches(0.6), Inches(0.3),
        Inches(12), Inches(0.7)
    )

    paragraph = box.text_frame.paragraphs[0]
    paragraph.text = title
    paragraph.font.size = Pt(28)
    paragraph.font.bold = True
    paragraph.font.color.rgb = DARK_BLUE


# Function to add body text
def add_text(slide, text, left, top, width, height, size=20):

    box = slide.shapes.add_textbox(
        Inches(left), Inches(top),
        Inches(width), Inches(height)
    )

    paragraph = box.text_frame.paragraphs[0]
    paragraph.text = text
    paragraph.font.size = Pt(size)
    paragraph.font.color.rgb = GREY


# ------------------------------------------
# SLIDE 1: INTRODUCTION
# ------------------------------------------

slide = prs.slides.add_slide(prs.slide_layouts[6])

add_text(
    slide,
    "NUTRITION PROGRAMME",
    1, 1.7, 11, 0.8, 34
)

add_text(
    slide,
    "PERFORMANCE ANALYSIS",
    1, 2.5, 11, 0.8, 34
)

add_text(
    slide,
    "A Data-Driven Review of Programme Outcomes",
    1, 3.5, 11, 0.6, 22
)

add_text(
    slide,
    "Client Presentation | November 2026",
    1, 4.5, 11, 0.5, 18
)

add_text(
    slide,
    "Prepared by: Data Analysis Team",
    1, 5.5, 11, 0.5, 16
)


# ------------------------------------------
# SLIDE 2: PROGRAMME OVERVIEW
# ------------------------------------------

slide = prs.slides.add_slide(prs.slide_layouts[6])

add_title(slide, "Programme Overview")

add_text(
    slide,
    f"Total Beneficiaries: {len(df)}",
    0.8, 1.7, 5.5, 0.8, 25
)

add_text(
    slide,
    f"Average Weight Gain: {df['weight_gain_kg'].mean():.2f} kg",
    0.8, 2.8, 6, 0.8, 25
)

add_text(
    slide,
    f"Average Sessions Attended: {df['sessions_attended'].mean():.1f}",
    0.8, 3.9, 7, 0.8, 25
)

add_text(
    slide,
    f"Counties Covered: {df['county'].nunique()}",
    0.8, 5.0, 6, 0.8, 25
)


# ------------------------------------------
# SLIDE 3: COUNTY PERFORMANCE
# ------------------------------------------

slide = prs.slides.add_slide(prs.slide_layouts[6])

add_title(slide, "Weight Gain Performance by County")

slide.shapes.add_picture(
    "charts/weight_gain_by_county.png",
    Inches(0.8), Inches(1.3),
    width=Inches(7.5)
)

best_county = county_analysis.idxmax()
best_county_value = county_analysis.max()

add_text(
    slide,
    f"Highest average weight gain:\n{best_county}\n\n"
    f"{best_county_value:.2f} kg",
    9, 2.0, 3.5, 2.5, 22
)


# ------------------------------------------
# SLIDE 4: KEY FINDINGS
# ------------------------------------------

slide = prs.slides.add_slide(prs.slide_layouts[6])

add_title(slide, "Key Findings: Age and Attendance")

slide.shapes.add_picture(
    "charts/weight_gain_by_age.png",
    Inches(0.3), Inches(1.4),
    width=Inches(6.0)
)

slide.shapes.add_picture(
    "charts/attendance_vs_weight_gain.png",
    Inches(6.7), Inches(1.4),
    width=Inches(6.0)
)

best_age = age_analysis.idxmax()
best_age_value = age_analysis.max()

add_text(
    slide,
    f"Highest average weight gain by age group: "
    f"{best_age} ({best_age_value:.2f} kg).",
    0.8, 6.3, 11.5, 0.6, 17
)


# ------------------------------------------
# SLIDE 5: RECOMMENDATIONS
# ------------------------------------------

slide = prs.slides.add_slide(prs.slide_layouts[6])

add_title(slide, "Recommendations and Next Steps")

recommendations = [
    "1. Strengthen attendance monitoring and follow-up.",
    "2. Investigate differences in outcomes across counties and age groups.",
    "3. Continue tracking weight changes to evaluate programme progress."
]

for i, recommendation in enumerate(recommendations):

    add_text(
        slide,
        recommendation,
        0.8, 1.7 + (i * 1.3),
        11.5, 0.8, 21
    )

add_text(
    slide,
    "Thank you | Questions and Discussion",
    0.8, 6.1, 11, 0.6, 22
)


# ==========================================
# 5. SAVE THE POWERPOINT
# ==========================================

prs.save("nutrition_programme_client_presentation.pptx")

print("\n================================")
print("POWERPOINT CREATED SUCCESSFULLY!")
print("Total slides:", len(prs.slides))
print("================================")
