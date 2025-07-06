import re

# Read the app_new.py file
with open('app_new.py', 'r') as f:
    content = f.read()

# Replace the first TOD summary section
pattern1 = r"""            # Add TOD-wise excess energy breakdown
            pdf\.ln\(5\)
            pdf\.set_font\('Arial', 'B', 12\)
            pdf\.cell\(0, 10, 'TOD-wise Excess Energy Breakdown:', ln=True\)
            pdf\.set_font\('Arial', 'B', 10\)
            pdf\.cell\(40, 8, 'TOD Category', 1\)
            pdf\.cell\(40, 8, 'Description', 1\)
            pdf\.cell\(40, 8, 'Excess Energy \(kWh\)', 1\)
            pdf\.ln\(\)
            pdf\.set_font\('Arial', '', 10\)
            
            # Get TOD-wise excess from the dataframe
            tod_excess = pdf_data\.groupby\('TOD_Category'\)\['Excess'\]\.sum\(\)\.reset_index\(\)
            
            tod_descriptions = \{
                'C1': 'Morning Peak',
                'C2': 'Evening Peak',
                'C4': 'Normal Hours',
                'C5': 'Night Hours',
                'Unknown': 'Unknown Time Slot'
            \}
            
            for _, row in tod_excess\.iterrows\(\):
                category = row\['TOD_Category'\]
                description = tod_descriptions\.get\(category, 'Unknown'\)
                excess = row\['Excess'\]
                pdf\.cell\(40, 8, category, 1\)
                pdf\.cell\(40, 8, description, 1\)
                pdf\.cell\(40, 8, f"\{excess:\\.4f\}", 1\)
                pdf\.ln\(\)"""

replacement1 = """            # Use the helper function to add TOD-wise excess energy breakdown
            tod_excess = add_tod_summary_table(pdf, pdf_data)"""

# Replace the second TOD summary section
pattern2 = r"""            # Add TOD-wise excess energy breakdown
            pdf\.ln\(5\)
            pdf\.set_font\('Arial', 'B', 12\)
            pdf\.cell\(0, 10, 'TOD-wise Excess Energy Breakdown:', ln=True\)
            pdf\.set_font\('Arial', 'B', 10\)
            pdf\.cell\(40, 8, 'TOD Category', 1\)
            pdf\.cell\(40, 8, 'Description', 1\)
            pdf\.cell\(40, 8, 'Excess Energy \(kWh\)', 1\)
            pdf\.ln\(\)
            pdf\.set_font\('Arial', '', 10\)
            
            # Get TOD-wise excess from the dataframe
            tod_excess = pdf_data\.groupby\('TOD_Category'\)\['Excess'\]\.sum\(\)\.reset_index\(\)
            
            tod_descriptions = \{
                'C1': 'Morning Peak',
                'C2': 'Evening Peak',
                'C4': 'Normal Hours',
                'C5': 'Night Hours',
                'Unknown': 'Unknown Time Slot'
            \}
            
            for _, row in tod_excess\.iterrows\(\):
                category = row\['TOD_Category'\]
                description = tod_descriptions\.get\(category, 'Unknown'\)
                excess = row\['Excess'\]
                pdf\.cell\(40, 8, category, 1\)
                pdf\.cell\(40, 8, description, 1\)
                pdf\.cell\(40, 8, f"\{excess:\\.4f\}", 1\)
                pdf\.ln\(\)"""

replacement2 = """            # Use the helper function to add TOD-wise excess energy breakdown
            tod_excess = add_tod_summary_table(pdf, pdf_data)"""

# Apply the replacements
modified_content = re.sub(pattern1, replacement1, content, count=1, flags=re.DOTALL)
modified_content = re.sub(pattern2, replacement2, modified_content, count=1, flags=re.DOTALL)

# Write the modified content to app_final.py
with open('app_final.py', 'w') as f:
    f.write(modified_content)

print("File updated successfully!")