import pandas as pd
import seaborn as sns
import matplotlib.pyplot as plt

def generate_ngx_market_cap_heatmap(pdf_filename='ngx_market_cap_heatmap.pdf'):
    """
    Generates a heatmap of the market capitalization for major companies
    listed on the Nigerian Exchange (NGX) and saves it as a PDF.
    
    The market capitalization data is pre-compiled based on publicly
    available information (in Trillion Nigerian Naira - T NGN).
    """
    
    # 1. Market Capitalization Data (Pre-compiled in T NGN - Trillion Naira)
    # This data is sourced from financial news and data providers on NGX companies.
    data = {
        'Company': ['BUAFOODS', 'MTNN', 'DANGCEM', 'BUACEMENT', 'GTCO', 'GEREGU', 
                    'ZENITHBANK', 'ARADEL', 'NB', 'INTBREW', 'TRANSPOWER', 'WAPCO', 
                    'UBA', 'STANBIC', 'TRANSCOHOT', 'NESTLE', 'PRESCO', 'ACCESSCORP', 
                    'FIRSTHOLDCO', 'FIDELITYBK'],
        'Market_Cap_TNGN': [11.33, 9.89, 9.54, 5.38, 3.46, 2.85, 2.81, 2.73, 2.42, 
                            2.36, 2.35, 2.09, 1.76, 1.73, 1.69, 1.48, 1.48, 1.36, 
                            1.30, 1.02]
    }
    
    df = pd.DataFrame(data)

    # 2. Data Preparation for Heatmap
    # Sort for better visualization and set 'Company' as index
    df_sorted = df.sort_values(by='Market_Cap_TNGN', ascending=False).set_index('Company')
    
    # Transpose the single-column DataFrame for a horizontal heatmap layout
    heatmap_data = df_sorted[['Market_Cap_TNGN']].T

    # 3. Plot Generation (Heatmap)
    plt.style.use('seaborn-v0_8-whitegrid') # Set a nice style
    plt.figure(figsize=(15, 6)) # Adjust size for many columns

    sns.heatmap(
        heatmap_data,
        annot=True,
        fmt=".2f", # Format annotation to 2 decimal places
        cmap="YlGnBu", # Color palette
        linewidths=.5,
        linecolor='black',
        cbar_kws={'label': 'Market Capitalization (T NGN)'},
        yticklabels=['Market Cap (T NGN)'] # Label for the single row
    )

    plt.title('Market Capitalization of Largest Companies on the Nigerian Exchange (NGX)', 
              fontsize=16, pad=20)
    plt.xlabel('Company Ticker Symbol', fontsize=12)
    plt.ylabel('')
    
    # Rotate x-axis labels for readability
    plt.xticks(rotation=45, ha='right')
    plt.yticks(rotation=0)
    plt.tight_layout() # Adjust plot to prevent labels from being cut off

    # 4. Save to PDF
    try:
        plt.savefig(pdf_filename, format='pdf')
        print("-" * 50)
        print(f"SUCCESS: Heatmap saved to {pdf_filename}")
        print(f"Top 5 Companies:")
        print(df_sorted.head(5).to_markdown(numalign="left", stralign="left"))
        print("-" * 50)
    except Exception as e:
        print(f"ERROR: Could not save the file. Ensure you have write permissions.")
        print(f"Details: {e}")
    finally:
        plt.close()

if __name__ == '__main__':
    # You may need to install the required libraries first:
    # pip install pandas seaborn matplotlib
    generate_ngx_market_cap_heatmap()

