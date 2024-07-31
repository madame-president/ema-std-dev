# ema-std-dev

Performs price forecasting and data analysis for a financial asset. It uses historical price data to create a simple forecast model based on the exponential moving average (EMA) and historical volatility.

## Features

1. **Data Loading and Preprocessing**
   - Reads price data from an Excel file
   - Converts dates to UTC format and sets them as the index
   - Calculates the percentage change in prices and the daily standard deviation

2. **Exponential Moving Average (EMA) Calculation**
   - Defines a function to calculate the EMA
   - Applies a 30-day EMA to the price data

3. **Price Forecasting**
   - Projects future prices using a function called `forecast_prices`
   - Uses the last known EMA and price as starting points
   - Generates random variations based on the historical daily standard deviation
   - Projects prices up to a specified end date (September 6, 2025)

4. **Data Combination and Formatting**
   - Combines historical data with forecasted data
   - Formats dates to match the input file format

5. **Excel Output**
   - Saves the combined data (historical + forecasted) to a new Excel file
   - Applies custom formatting to the Excel file:
     - Sets a specific date format for the 'Date' column
     - Applies currency formatting to price columns

6. **Final Output**
   - Prints confirmation messages about:
     - Forecast generation
     - File saving
     - Daily standard deviation of price changes

## Usage

1. Ensure you have the required libraries installed: pandas, numpy, pytz, openpyxl
2. Prepare your input Excel file with historical price data
3. Update the `price_input` variable with your input file name
4. Set the `output_file` variable to your desired output file name
5. Run the script

## Note

The forecast extends to September 6, 2025.

## Customization

You can adjust the following parameters in the script:
- EMA period (currently set to 30 days)
- Forecast end date
- Input and output file names
