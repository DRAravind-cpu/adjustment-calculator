import pandas as pd
from datetime import datetime

def compute_additional_surcharge(pdf_data):
    additional_surcharge = 0.0
    additional_surcharge_breakdown = []
    try:
        slot_dates_dt = pd.to_datetime(pdf_data['Slot_Date'], errors='coerce', dayfirst=True)
        surcharge_ranges = [
            ('2021-04-16', '2022-03-31', 0.70, "TNERC M.P.No.18 of 2020 dt.15.04.2021"),
            ('2023-02-25', '2023-03-31', 0.83, "TNERC M.P.No.32 of 2021 dt.08.02.2022"),
            ('2024-12-12', '2025-03-31', 0.54, "TNERC M.P.No.44 of 2024 dt.12.12.2024"),
            ('2025-04-29', '2025-09-30', 0.10, "TNERC M.P.No.13 of 2025 dt.29.04.2025")
        ]
        for start_str, end_str, rate, note in surcharge_ranges:
            start_dt = pd.to_datetime(start_str)
            end_dt = pd.to_datetime(end_str)
            mask = (slot_dates_dt >= start_dt) & (slot_dates_dt <= end_dt)
            if mask.any() and 'IEX_Excess' in pdf_data.columns:
                iex_sum_raw = pdf_data.loc[mask, 'IEX_Excess'].sum()
                iex_sum = int(iex_sum_raw + 0.5) if iex_sum_raw >= 0 else int(iex_sum_raw - 0.5)
                component = iex_sum * rate
                additional_surcharge += component
                additional_surcharge_breakdown.append((start_str, end_str, iex_sum, rate, component, note))
    except Exception as e:
        print('Error', e)
    return additional_surcharge, additional_surcharge_breakdown

# Create sample data spanning two ranges
data = {
    'Slot_Date': ['16/04/2021', '17/04/2021', '01/03/2023', '26/02/2023', '15/12/2024'],
    'IEX_Excess': [1000, 500, 200, 800, 1200]
}
pdf_data = pd.DataFrame(data)
print(pdf_data)
add, breakdown = compute_additional_surcharge(pdf_data)
print('Additional surcharge:', add)
print('Breakdown:', breakdown)
