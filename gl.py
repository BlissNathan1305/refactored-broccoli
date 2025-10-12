import pandas as pd
import matplotlib.pyplot as plt
from fpdf import FPDF

# Sample data based on 2025 reports
data = {
    'Usage Category': ['Communication', 'Entertainment', 'Productivity', 'Health Tracking', 'Shopping'],
    'Percentage of Users': [95, 88, 76, 64, 70]
}

# Create DataFrame
df = pd.DataFrame(data)

# Basic statistics
summary = df.describe()

# Plotting
plt.figure(figsize=(10, 6))
plt.bar(df['Usage Category'], df['Percentage of Users'], color='skyblue')
plt.title('Mobile Device Usage by Category (2025)')
plt.xlabel('Category')
plt.ylabel('Percentage of Users')
plt.tight_layout()
plt.savefig('usage_chart.png')
plt.close()

# PDF Export
pdf = FPDF()
pdf.add_page()
pdf.set_font("Arial", size=12)
pdf.cell(200, 10, txt="Mobile Device Usage Analysis - 2025", ln=True, align='C')

# Add summary stats
pdf.ln(10)
pdf.cell(200, 10, txt="Statistical Summary:", ln=True)
for index, row in summary.iterrows():
    pdf.cell(200, 10, txt=f"{index}: {row['Percentage of Users']:.2f}", ln=True)

# Add chart
pdf.image('usage_chart.png', x=10, y=80, w=190)

# Save PDF
pdf.output("Mobile_Usage_Analysis_2025.pdf")
