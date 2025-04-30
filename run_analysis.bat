@echo off
REM Batch file to run the time series analysis with specific parameters

echo Running Time Series Analysis...

python main.py "C:\Users\franc\OneDrive\Desktop\Programming\Regression Project\fredgraph.xlsx" "date" "icsa" "unrate" --range 12 --window 12 --header 0 --sheet Monthly --output_dir "C:\Users\franc\OneDrive\Desktop\Programming\Regression Project"

echo.
echo Analysis script finished.
pause
