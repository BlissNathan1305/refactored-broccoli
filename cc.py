from reportlab.lib.pagesizes import letter
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, PageBreak, Preformatted
from reportlab.lib.styles import getSampleStyleSheet
from reportlab.lib.units import inch
from reportlab.lib import colors

# --- Document Metadata ---
book_title = "Using Python and R for Data Analysis"
subtitle = "A practical guide with examples"
author = "Author: Data Analyst"

# --- Content Definition ---
# Refactored content_paragraphs using proper string formatting (triple quotes for code/blocks)
# and clearly defining block elements (Heading, Code, Normal).

# Define content as a list of tuples: (content_type, content)
# content_type: 'H1' for Heading1, 'H2' for Heading2 (for subtitle), 'P' for Normal Paragraph, 'CODE' for Preformatted Code/Blockquote
content_blocks = [
    # Cover page elements (will be handled separately, but included for clarity)
    ('TITLE', book_title),
    ('H2', subtitle),
    ('AUTHOR', author),

    # Main Content
    ('H1', "Introduction"),
    ('P',
        "Data analysis is the process of inspecting, cleansing, transforming, and modeling data to extract useful information, "
        "inform conclusions, and support decision-making. Python and R are two powerful programming languages widely used in data analysis."
    ),
    ('P',
        "Python is known for its versatility, extensive libraries, and integration capabilities, making it popular for general data analysis, "
        "machine learning, and web applications. R is a statistical programming language tailored for deep statistical analysis and high-quality data visualization."
    ),
    ('P',
        "This book covers the fundamental types of data analysis, how they can be performed using Python and R, and provides practical code examples "
        "for each type. Additional chapters discuss advanced topics such as data visualization, cleaning, and machine learning."
    ),

    ('H1', "Types of Data Analysis"),
    ('P',
        "1. Descriptive Analysis: Summarizes key characteristics of data using statistics and visualizations.\n"
        "2. Exploratory Data Analysis (EDA): Analyzes data sets to find patterns, anomalies, and test hypotheses visually and statistically.\n"
        "3. Inferential Analysis: Draws conclusions about a population using data from a sample, involving hypothesis testing and confidence intervals.\n"
        "4. Predictive Analysis: Uses historical data to build models that predict future outcomes.\n"
        "5. Causal Analysis: Identifies cause-effect relationships within the data.\n"
        "6. Mechanistic Analysis: Investigates underlying mechanisms or processes generating the data.\n"
        "7. Diagnostic Analysis: Examines data to understand why something happened.\n"
        "8. Prescriptive Analysis: Suggests possible actions based on data insights to achieve desired outcomes."
    ),

    ('H1', "Python for Data Analysis"),
    ('P',
        "Python’s rich ecosystem includes:"
    ),
    ('P',
        "• pandas: Data manipulation and analysis\n"
        "• matplotlib & seaborn: Visualization libraries\n"
        "• scikit-learn: Machine learning\n"
        "• statsmodels: Statistical modeling"
    ),
    ('P', "Python supports multiple data formats and scales well for large datasets."),
    ('P', "Example: Descriptive Analysis in Python"),
    ('CODE',
"""import pandas as pd
import matplotlib.pyplot as plt

data = {
    'Age': [23, 45, 31, 35, 50],
    'Income': [50000, 80000, 60000, 65000, 90000]
}

df = pd.DataFrame(data)

# Summary statistics
print(df.describe())

# Histogram of Age
df['Age'].hist()
plt.title('Age Distribution')
plt.xlabel('Age')
plt.ylabel('Frequency')
plt.show()"""),

    ('H1', "R for Data Analysis"),
    ('P',
        "R excels in statistical computing and graphics. Commonly used packages:\n"
        "• ggplot2: Data visualization\n"
        "• dplyr: Data manipulation\n"
        "• caret: Machine learning\n"
        "• shiny: Interactive web apps"
    ),
    ('P', "Its syntax is specialized for statistical tasks and it integrates well with statistical tests."),
    ('P', "Example: Descriptive Analysis in R"),
    ('CODE',
"""data <- data.frame(
  Age = c(23, 45, 31, 35, 50),
  Income = c(50000, 80000, 60000, 65000, 90000)
)

# Summary statistics
summary(data)

# Histogram of Age
hist(data$Age, main = "Age Distribution", xlab = "Age", ylab = "Frequency")"""),

    ('H1', "Exploratory Data Analysis (EDA)"),
    ('P', "EDA helps to understand data distributions, spot outliers, and check assumptions."),
    ('P', "Python example using seaborn:"),
    ('CODE',
"""import seaborn as sns

sns.pairplot(df)
plt.show()"""),
    ('P', "R example using ggplot2:"),
    ('CODE',
"""library(ggplot2)

ggplot(data, aes(x = Age, y = Income)) +
  geom_point() +
  geom_smooth(method = "lm") +
  ggtitle("Age vs Income")"""),

    ('H1', "Inferential Statistics"),
    ('P', "Conduct hypothesis tests and create confidence intervals to infer about populations."),
    ('P', "Python example using scipy:"),
    ('CODE',
"""from scipy import stats

# T-test for mean Age vs hypothetical mean 30
t_stat, p_val = stats.ttest_1samp(df['Age'], 30)
print(f'T-statistic: {t_stat}, P-value: {p_val}')"""),
    ('P', "R example:"),
    ('CODE', "t.test(data$Age, mu = 30)"),

    ('H1', "Predictive Modeling"),
    ('P', "Python example with linear regression:"),
    ('CODE',
"""from sklearn.linear_model import LinearRegression
import numpy as np

X = df[['Age']]
y = df['Income']

model = LinearRegression()
model.fit(X, y)
print(f'Coefficients: {model.coef_}, Intercept: {model.intercept_}')"""),
    ('P', "R example:"),
    ('CODE',
"""model <- lm(Income ~ Age, data = data)
summary(model)"""),

    ('H1', "Causal and Mechanistic Analysis"),
    ('P', "These analyses identify causes or mechanisms behind observed data and often require experimental or longitudinal data."),

    ('H1', "Data Visualization Best Practices"),
    ('P',
        "• Use clear labels and legends.\n"
        "• Choose appropriate chart types.\n"
        "• Avoid chart junk and excessive colors.\n"
        "• Use visualization libraries’ built-in themes for consistency."
    ),

    ('H1', "Data Cleaning and Preparation"),
    ('P', "Process includes handling missing data, correcting data types, removing duplicates, and feature engineering."),

    ('H1', "Working with Big Data"),
    ('P', "Python libraries like Dask and R packages such as data.table enable big data processing."),

    ('H1', "Time Series Analysis"),
    ('P', "Focuses on ordered data over time for trend and seasonality analysis."),

    ('H1', "Text Data Analysis / NLP"),
    ('P', "Use Python libraries like NLTK and R’s tm package for processing text data."),

    ('H1', "Machine Learning Overview"),
    ('P', "Python’s scikit-learn and R’s caret simplify building classification, regression, and clustering models."),

    ('H1', "Case Studies"),
    ('P', "Real-world data projects demonstrating end-to-end workflows in Python and R."),

    ('H1', "Closing Remarks"),
    ('P', "This book equips readers with fundamentals and practical skills in data analysis using Python and R. Practice with the included examples to develop proficiency.")
]


def create_pdf(filename):
    doc = SimpleDocTemplate(
        filename,
        pagesize=letter,
        rightMargin=72,
        leftMargin=72,
        topMargin=72,
        bottomMargin=72
    )
    styles = getSampleStyleSheet()
    normal_style = styles['Normal']
    title_style = styles['Title']
    heading1_style = styles['Heading1']
    heading2_style = styles['Heading2']
    italic_style = styles['Italic']

    # Custom style for code blocks (monospaced font, smaller size, grey background)
    code_style = styles['Code']
    code_style.fontSize = 8
    code_style.textColor = colors.black
    code_style.backColor = colors.lightgrey

    story = []

    # --- Cover page ---
    story.append(Paragraph(book_title, title_style))
    story.append(Spacer(1, 12))
    story.append(Paragraph(subtitle, heading2_style))
    story.append(Spacer(1, 12))
    story.append(Paragraph(author, italic_style))
    story.append(PageBreak())

    # --- Content Loop ---
    for content_type, content in content_blocks:
        if content_type == 'H1':
            story.append(Paragraph(content, heading1_style))
            story.append(Spacer(1, 6)) # Less space after a heading
        elif content_type == 'P':
            # Replace single newlines within P blocks with <br/> for ReportLab to handle
            # This is necessary for things like list items within a single 'P' block
            formatted_content = content.replace('\n', '<br/>')
            story.append(Paragraph(formatted_content, normal_style))
            story.append(Spacer(1, 12))
        elif content_type == 'CODE':
            # Use Preformatted for code blocks to preserve newlines and indentation
            # ReportLab's Paragraph element is not meant for code blocks.
            story.append(Preformatted(content, code_style))
            story.append(Spacer(1, 12))

    doc.build(story)
    print(f"PDF '{filename}' created successfully.")

if __name__ == "__main__":
    create_pdf("Python_R_Data_Analysis_Book.pdf")

