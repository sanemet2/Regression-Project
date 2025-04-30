# Project Plan: Time Series Lead/Lag Analysis App

## 1. Goal

Create a Python Command Line Interface (CLI) application that analyzes the lead/lag relationship between two time series from an Excel file. The app will identify the optimal lead/lag time based on R-squared, visualize the relationship, calculate rolling correlations for different leads/lags, and export results to Excel.

## 2. Core Features

*   **Data Loading:**
    *   Load data from a user-specified Excel file (`.xlsx` or `.xls`).
    *   Allow the user to specify:
        *   Date column name.
        *   Leading series column name.
        *   Target series column name.
    *   Parse dates assuming `mm/yy` format.
    *   Handle missing data by dropping rows with `NaN` in selected columns.
*   **Optimal Lead/Lag Analysis:**
    *   User specifies the maximum lead/lag period `N` (e.g., 12).
    *   The application tests shifts from `-N` to `+N` periods.
    *   For each shift, align the series and calculate the R-squared value between the shifted leading series and the target series.
    *   Identify the shift (lead/lag) with the highest R-squared.
    *   Output:
        *   Print the optimal lead/lag period and the corresponding R-squared value.
        *   Generate and save a scatter plot of the optimally shifted leading series vs. the target series.
        *   Generate and save a line chart showing the target series and the optimally shifted leading series.
*   **Rolling Correlation Analysis:**
    *   User specifies the rolling window size (e.g., 36 periods).
    *   For each lead/lag tested (`-N` to `+N`), calculate the rolling correlation between the shifted leading series and the target series using the specified window.
    *   Output:
        *   Generate and save a line chart plotting the rolling correlations for *all* tested leads/lags over time.
*   **Export Results:**
    *   Create an output Excel file (`results.xlsx`).
    *   Include sheets for:
        *   `LeadLag_R2`: Table showing each tested lead/lag and its R-squared value.
        *   `Optimal_Shift_Data`: Columns for Date, Target Series, Leading Series (optimally shifted).
        *   `Rolling_Correlations`: Columns for Date and the rolling correlation values for each tested lead/lag.
    *   Save generated plots as separate image files (e.g., `optimal_scatter.png`, `optimal_line.png`, `rolling_corr.png`).

## 3. Technical Stack

*   **Language:** Python 3.x
*   **Core Libraries:**
    *   `pandas`: Data loading, manipulation, date handling, rolling calculations.
    *   `numpy`: Numerical operations.
    *   `scikit-learn`: Calculating R-squared (`sklearn.metrics.r2_score`).
    *   `matplotlib` / `seaborn`: Generating plots.
    *   `openpyxl` / `xlsxwriter`: Reading/writing Excel files.
    *   `argparse`: Handling command-line arguments.

## 4. Project Structure (Initial)

```
Stats/
|-- main.py             # Main script, argument parsing, orchestration
|-- analysis.py         # Core analysis functions (lead/lag, rolling corr)
|-- data_loader.py      # Function to load and preprocess data
|-- plotting.py         # Functions for generating plots
|-- export.py           # Functions for exporting results to Excel
|-- requirements.txt    # Project dependencies
|-- PLAN.md             # This plan file
|-- results/            # Directory to save output plots and Excel file (created by app)
```

## 5. Execution Steps (Sequential)

- [x] **Setup:** Create project structure, `requirements.txt`.
- [x] **Data Loading (`data_loader.py`):** Implement function to load Excel, select columns, parse dates, handle NaNs.
- [x] **CLI (`main.py`):** Set up `argparse` to take file path, column names, lead/lag range (`N`), and rolling window size as input.
- [x] **Lead/Lag Analysis (`analysis.py`):** Implement function to iterate through shifts (-N to +N), calculate R-squared for each, return best shift and R-squared values.
- [x] **Plotting (`plotting.py`):
    - [x] Implement function for scatter plot.
    - [x] Implement function for optimal lead line chart.
- [x] **Rolling Correlation (`analysis.py`):** Implement function to calculate rolling correlations for all shifts.
- [x] **Plotting (`plotting.py`):** Implement function for rolling correlation evolution chart.
- [x] **Export (`export.py`):** Implement function to write results (R-squared table, shifted data, rolling correlations) to `results.xlsx`.
- [x] **Integration (`main.py`):** Connect all parts: Load data, run analyses, generate plots (saving them), export results.
- [x] **Testing:** Test with sample data, refine as needed.

## 6. Future Considerations

*   GUI Development (e.g., using Tkinter, PyQt, or Streamlit).
*   More sophisticated missing data handling options.
*   Additional statistical tests (e.g., p-values for correlations).
*   Embedding plots directly into Excel (using `xlsxwriter` capabilities).
*   Winsorizing / trimming outliers
*   Automatically adjust time periods if one is weekly, the other monthly etc.
*   Fix charting functions and rolling correlation charts
