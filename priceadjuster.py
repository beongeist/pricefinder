import pandas as pd
import math

# Read the Excel file into a pandas DataFrame
df = pd.read_excel("PSA inventory.xlsx")

# Function to clean and convert price strings to numeric values
def clean_price(price_str):
    # Remove '$' and ',' from the price string, then convert to float
    try:
        return float(price_str.replace('$', '').replace(',', ''))
    except:
        print("fail")

def adjust_price(price):
    if pd.isna(price):  # Skip NaN values
        return price

    # Round to the nearest dollar
    rounded_price = round(price)

    # Determine rounding criteria
    if rounded_price % 5 == 0:  # Already a multiple of 5
        return rounded_price
    elif (rounded_price + 1) % 5 == 0 or (rounded_price + 2) % 5 == 0:  # Rounds up to a multiple of 5
        return math.ceil(price / 5) * 5
    elif (rounded_price - 1) % 10 == 0 or (rounded_price - 2) % 10 == 0:  # Rounds down to a multiple of 10
        return math.floor(price / 10) * 10
    else:
        # Round up to the nearest dollar
        return math.ceil(price)

# Apply the function to the 'price' column and create a new column with adjusted prices

df['My Cost'] = df['Purchase Date'].apply(adjust_price)

df.to_excel("PSA inventory.xlsx", index=False)
