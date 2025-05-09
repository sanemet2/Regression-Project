import streamlit as st
import pandas as pd
import numpy as np
from data_loader import load_data
from analysis import find_optimal_lead_lag, calculate_rolling_correlations, calculate_cumulative_correlations
from export import export_to_excel as export_results

# Original function to highlight rows with max Cumulative_R2 and max Rolling_R2
def highlight_max_r_squared_rows(df_to_style, columns_to_highlight):
    # Create a DataFrame of styles, initialized to no style
    styler_df = pd.DataFrame('', index=df_to_style.index, columns=df_to_style.columns)
    highlight_color_css = 'background-color: lightgrey; color: black;'

    # --- Highlight for Max Cumulative_R2 ---
    cumulative_r2_numeric = pd.to_numeric(df_to_style['Cumulative_R2'], errors='coerce')
    if cumulative_r2_numeric.notna().any(): 
        max_cum_idx = cumulative_r2_numeric.idxmax()
        if pd.notna(max_cum_idx): 
            for col_name in columns_to_highlight:
                if col_name in styler_df.columns: # Check if column exists
                    styler_df.loc[max_cum_idx, col_name] = highlight_color_css

    # --- Highlight for Max Rolling_R2 ---
    rolling_r2_numeric = pd.to_numeric(df_to_style['Rolling_R2'], errors='coerce')
    if rolling_r2_numeric.notna().any(): 
        max_roll_idx = rolling_r2_numeric.idxmax()
        if pd.notna(max_roll_idx): 
            for col_name in columns_to_highlight:
                if col_name in styler_df.columns: # Check if column exists
                    styler_df.loc[max_roll_idx, col_name] = highlight_color_css

    return styler_df

# Main entry point for the Streamlit app
def main():
    st.markdown("<h1 style='color: red;'>Mixed-Frequency Analysis GUI</h1>", unsafe_allow_html=True)

    # Restore Step 3: File uploader and sheet selector with error handling
    uploaded_file = st.file_uploader("Upload Excel workbook", type=["xls", "xlsx"])
    if uploaded_file is not None:
        try:
            # Use 'with' statement for proper file handling
            with pd.ExcelFile(uploaded_file) as xls:
                sheet_names = xls.sheet_names
                st.write("**Available sheets:**", sheet_names)

                sheet = st.selectbox("Select sheet", sheet_names)

                if sheet:
                    try:
                        # Step 4: Load data from selected sheet
                        df_raw = xls.parse(sheet)
                        st.write(f"**Preview of {sheet}:**")
                        st.dataframe(df_raw)

                        # Step 5: Dynamic Column Selection
                        if not df_raw.empty:
                            columns = df_raw.columns.tolist()
                            st.subheader("Select Columns for Analysis")
                            col1, col2 = st.columns(2)
                            with col1:
                                lead_date_col = st.selectbox("Lead Date Column", columns)
                                lead_val_col = st.selectbox("Lead Value Column", columns)
                            with col2:
                                target_date_col = st.selectbox("Static Date Column", columns)
                                target_val_col = st.selectbox("Static Value Column", columns)

                        # Step 6: Analysis Parameter Widgets
                        st.subheader("Analysis Parameters")
                        col3, col4 = st.columns(2)
                        with col3:
                            max_shift = st.number_input("Max Lead/Lag Shift (+/- periods)", min_value=1, value=12, step=1)
                            window = st.number_input("Rolling Correlation Window (periods)", min_value=2, value=12, step=1)
                        with col4:
                             exclude_period = st.text_input("Exclude Period (YYYY-MM:YYYY-MM or START:END)", "")

                        # Step 7: Run Analysis Button and Logic
                        if st.button("Run Analysis"):
                            st.write("--- Running Analysis ---")
                            # Ensure we have the uploaded file object for load_data
                            if uploaded_file is not None:
                                try:
                                    # 7.a: Call load_data with selected inputs
                                    # Need header_row - assuming 0 for now, might need to be configurable?
                                    # Also need the actual file path or buffer
                                    # uploaded_file itself works as a buffer for pandas
                                    st.write(f"Calling load_data with:")
                                    st.write(f"  Sheet: {sheet}")
                                    st.write(f"  Lead Date: {lead_date_col}, Lead Val: {lead_val_col}")
                                    st.write(f"  Static Date: {target_date_col}, Static Val: {target_val_col}")

                                    lead_series, target_series = load_data(
                                        file_path=uploaded_file, # Pass the buffer directly
                                        sheet_name=sheet,
                                        lead_date_col=lead_date_col,
                                        lead_val_col=lead_val_col,
                                        target_date_col=target_date_col,
                                        target_val_col=target_val_col,
                                        header_row=0 # TODO: Make header row configurable?
                                    )

                                    if lead_series is not None and target_series is not None:
                                        st.success("Data loaded successfully!")

                                        # Step 7.b: Apply Exclude Period
                                        if exclude_period:
                                            try:
                                                start_str, end_str = exclude_period.split(':')
                                                # Attempt to parse as YYYY-MM first, then fallback if needed
                                                start_date = pd.to_datetime(start_str, format='%Y-%m', errors='coerce')
                                                end_date = pd.to_datetime(end_str, format='%Y-%m', errors='coerce')
                                                
                                                # More robust parsing if YYYY-MM fails (e.g., full dates)
                                                if pd.isna(start_date):
                                                     start_date = pd.to_datetime(start_str, errors='coerce')
                                                if pd.isna(end_date):
                                                     end_date = pd.to_datetime(end_str, errors='coerce')

                                                if pd.isna(start_date) or pd.isna(end_date):
                                                    st.warning(f"Could not parse exclude period '{exclude_period}'. Please use format YYYY-MM:YYYY-MM or a pandas-parsable date format.")
                                                else:
                                                    st.write(f"Applying exclusion period: {start_date.strftime('%Y-%m-%d')} to {end_date.strftime('%Y-%m-%d')}")
                                                    lead_rows_before = len(lead_series)
                                                    target_rows_before = len(target_series)

                                                    lead_series = lead_series[~((lead_series.index >= start_date) & (lead_series.index <= end_date))]
                                                    target_series = target_series[~((target_series.index >= start_date) & (target_series.index <= end_date))]

                                                    st.write(f"  - Lead series rows removed: {lead_rows_before - len(lead_series)}")
                                                    st.write(f"  - Static series rows removed: {target_rows_before - len(target_series)}")
                                            except ValueError:
                                                st.warning(f"Invalid format for exclude period '{exclude_period}'. Use START:END format, e.g., 2020-03:2020-06.")
                                            except Exception as ex_err:
                                                st.warning(f"Error applying exclusion period: {ex_err}")

                                        # Step 8: Frequency Detection, Resampling, and Alignment
                                        st.write("--- Preparing Data for Analysis ---")
                                        try:
                                            freq_lead = pd.infer_freq(lead_series.index)
                                            freq_target = pd.infer_freq(target_series.index)
                                            st.write(f"Detected Frequencies -> Lead: {freq_lead}, Static: {freq_target}")

                                            resampled = False
                                            # Simple check: if frequencies differ and are detected
                                            if freq_lead and freq_target and freq_lead != freq_target:
                                                # Determine which is higher freq
                                                lead_period = pd.tseries.frequencies.to_offset(freq_lead)
                                                target_period = pd.tseries.frequencies.to_offset(freq_target)

                                                # Normalize base codes and compare frequency hierarchy
                                                freq_order = {'D': 1, 'B': 2, 'W': 3, 'M': 4, 'Q': 5, 'A': 6}
                                                def get_base_code(freq_str):
                                                    base = freq_str.split('-')[0]
                                                    return base.rstrip('S')

                                                base_lead = get_base_code(freq_lead)
                                                base_target = get_base_code(freq_target)
                                                rank_lead = freq_order.get(base_lead)
                                                rank_target = freq_order.get(base_target)

                                                if rank_lead is not None and rank_target is not None:
                                                    if rank_lead < rank_target: # lead is higher frequency
                                                        st.write(f"Resampling Lead series ({freq_lead}) to match Static series ({freq_target})...")
                                                        lead_series = lead_series.resample(target_period).last()
                                                        resampled = True
                                                        freq_lead = freq_target
                                                    elif rank_target < rank_lead: # target is higher frequency
                                                        st.write(f"Resampling Static series ({freq_target}) to match Lead series ({freq_lead})...")
                                                        target_series = target_series.resample(lead_period).last()
                                                        resampled = True
                                                        freq_target = freq_lead
                                                else:
                                                    st.warning(f"Could not determine frequency hierarchy between '{freq_lead}' and '{freq_target}'. Skipping resampling.")

                                            # Combine into a DataFrame, aligning by index
                                            st.write(f"Aligning series at frequency: {freq_lead or freq_target or 'Undetected'}")
                                            analysis_df = pd.DataFrame({'Leading': lead_series, 'Target': target_series})

                                            # Drop rows with NaNs resulting from alignment/resampling
                                            rows_before_na_drop = len(analysis_df)
                                            analysis_df.dropna(inplace=True)
                                            rows_after_na_drop = len(analysis_df)

                                            if rows_before_na_drop > rows_after_na_drop:
                                                st.write(f"Dropped {rows_before_na_drop - rows_after_na_drop} rows with missing values after alignment/resampling.")

                                            if analysis_df.empty:
                                                st.error("No overlapping data remains after alignment and NA removal. Cannot proceed with analysis.")
                                                # Consider setting a flag here to prevent further steps
                                            else:
                                                st.success("Data successfully aligned and prepared.")
                                                st.write("**Aligned Data Preview:**")
                                                st.dataframe(analysis_df.head())

                                                print("--- Debug: analysis_df dtypes before analysis --- ")
                                                print(analysis_df.info())

                                                # Run analysis functions and compute metrics
                                                try:
                                                    rolling_df = calculate_rolling_correlations(analysis_df, max_shift, window)
                                                    cum_df = calculate_cumulative_correlations(analysis_df, max_shift)
                                                    # Summary table of final R-squared by shift
                                                    shifts = list(range(-max_shift, max_shift + 1))
                                                    cum_last = cum_df.iloc[-1]
                                                    roll_last = rolling_df.iloc[-1]
                                                    cum_vals = [cum_last.get(f'CumCorr_Shift_{s}', None) for s in shifts]
                                                    roll_vals = [roll_last.get(f'Shift_{s}', None) for s in shifts]
                                                    summary_df = pd.DataFrame({
                                                        'Shift': shifts,
                                                        'Cumulative_R2': cum_vals,
                                                        'Rolling_R2': roll_vals
                                                    })

                                                    # Apply custom highlighting
                                                    columns_to_highlight = ['Shift', 'Cumulative_R2', 'Rolling_R2'] # Target the actual 'Shift' data column
                                                    styled_summary_df = summary_df.style.apply(highlight_max_r_squared_rows, columns_to_highlight=columns_to_highlight, axis=None)
                                                    st.subheader("Summary of Final R-Squared by Shift")
                                                    st.dataframe(styled_summary_df)
                                                except Exception as e:
                                                    st.error(f"An error occurred during analysis: {e}")
                                        except Exception as align_error:
                                            st.error("An error occurred during frequency detection or alignment:")
                                            st.exception(align_error)

                                    else:
                                        st.error("Failed to load data. Check column selections and sheet format.")

                                except Exception as load_error:
                                    st.error("An error occurred during data loading:")
                                    st.exception(load_error)
                            else:
                                st.warning("Cannot run analysis, file upload object is missing.")

                    except Exception as e:
                        st.error(f"Error parsing sheet '{sheet}' or setting up widgets:")
                        st.exception(e)

        except Exception as e:
            st.error("Error processing Excel file:")
            st.exception(e)

if __name__ == "__main__":
    main()
