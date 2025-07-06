from flask import Flask, render_template, request, send_file
import pandas as pd
from fpdf import FPDF
import io
import os

# Helper function to add TOD summary table with consistent formatting
def add_tod_summary_table(pdf, pdf_data):
    # Define TOD descriptions
    tod_descriptions = {
        'C1': 'Morning Peak',
        'C2': 'Evening Peak',
        'C4': 'Normal Hours',
        'C5': 'Night Hours',
        'Unknown': 'Unknown Time Slot'
    }
    
    # Add TOD-wise excess energy breakdown
    pdf.ln(5)
    pdf.set_font('Arial', 'B', 12)
    pdf.cell(0, 10, 'TOD-wise Excess Energy Breakdown:', ln=True)
    pdf.set_font('Arial', 'B', 10)
    pdf.cell(30, 8, 'TOD Category', 1)
    pdf.cell(80, 8, 'Description', 1)
    pdf.cell(40, 8, 'Excess Energy (kWh)', 1)
    pdf.ln()
    pdf.set_font('Arial', '', 10)
    
    # Get TOD-wise excess from the dataframe
    tod_excess = pdf_data.groupby('TOD_Category')['Excess'].sum().reset_index()
    
    # Sort by TOD category for consistent display
    tod_excess = tod_excess.sort_values('TOD_Category')
    
    for _, row in tod_excess.iterrows():
        category = row['TOD_Category']
        description = tod_descriptions.get(category, 'Unknown')
        excess = row['Excess']
        pdf.cell(30, 8, category, 1)
        pdf.cell(80, 8, description, 1)
        pdf.cell(40, 8, f"{excess:.4f}", 1)
        pdf.ln()
    
    return tod_excessfrom flask import Flask, render_template, request, send_file
import pandas as pd
from fpdf import FPDF
import io
import os

# Helper function to add TOD summary table with consistent formatting
def add_tod_summary_table(pdf, pdf_data):
    # Define TOD descriptions
    tod_descriptions = {
        'C1': 'Morning Peak',
        'C2': 'Evening Peak',
        'C4': 'Normal Hours',
        'C5': 'Night Hours',
        'Unknown': 'Unknown Time Slot'
    }
    
    # Add TOD-wise excess energy breakdown
    pdf.ln(5)
    pdf.set_font('Arial', 'B', 12)
    pdf.cell(0, 10, 'TOD-wise Excess Energy Breakdown:', ln=True)
    pdf.set_font('Arial', 'B', 10)
    pdf.cell(30, 8, 'TOD Category', 1)
    pdf.cell(80, 8, 'Description', 1)
    pdf.cell(40, 8, 'Excess Energy (kWh)', 1)
    pdf.ln()
    pdf.set_font('Arial', '', 10)
    
    # Get TOD-wise excess from the dataframe
    tod_excess = pdf_data.groupby('TOD_Category')['Excess'].sum().reset_index()
    
    # Sort by TOD category for consistent display
    tod_excess = tod_excess.sort_values('TOD_Category')
    
    for _, row in tod_excess.iterrows():
        category = row['TOD_Category']
        description = tod_descriptions.get(category, 'Unknown')
        excess = row['Excess']
        pdf.cell(30, 8, category, 1)
        pdf.cell(80, 8, description, 1)
        pdf.cell(40, 8, f"{excess:.4f}", 1)
        pdf.ln()
    
    return tod_excessfrom flask import Flask, render_template, request, send_file
import pandas as pd
from fpdf import FPDF
import io
import os

# Helper function to add TOD summary table with consistent formatting
def add_tod_summary_table(pdf, pdf_data):
    # Define TOD descriptions
    tod_descriptions = {
        'C1': 'Morning Peak',
        'C2': 'Evening Peak',
        'C4': 'Normal Hours',
        'C5': 'Night Hours',
        'Unknown': 'Unknown Time Slot'
    }
    
    # Add TOD-wise excess energy breakdown
    pdf.ln(5)
    pdf.set_font('Arial', 'B', 12)
    pdf.cell(0, 10, 'TOD-wise Excess Energy Breakdown:', ln=True)
    pdf.set_font('Arial', 'B', 10)
    pdf.cell(30, 8, 'TOD Category', 1)
    pdf.cell(80, 8, 'Description', 1)
    pdf.cell(40, 8, 'Excess Energy (kWh)', 1)
    pdf.ln()
    pdf.set_font('Arial', '', 10)
    
    # Get TOD-wise excess from the dataframe
    tod_excess = pdf_data.groupby('TOD_Category')['Excess'].sum().reset_index()
    
    # Sort by TOD category for consistent display
    tod_excess = tod_excess.sort_values('TOD_Category')
    
    for _, row in tod_excess.iterrows():
        category = row['TOD_Category']
        description = tod_descriptions.get(category, 'Unknown')
        excess = row['Excess']
        pdf.cell(30, 8, category, 1)
        pdf.cell(80, 8, description, 1)
        pdf.cell(40, 8, f"{excess:.4f}", 1)
        pdf.ln()
    
    return tod_excessfrom flask import Flask, render_template, request, send_file
import pandas as pd
from fpdf import FPDF
import io
import os

# Helper function to add TOD summary table with consistent formatting
def add_tod_summary_table(pdf, pdf_data):
    # Define TOD descriptions
    tod_descriptions = {
        'C1': 'Morning Peak',
        'C2': 'Evening Peak',
        'C4': 'Normal Hours',
        'C5': 'Night Hours',
        'Unknown': 'Unknown Time Slot'
    }
    
    # Add TOD-wise excess energy breakdown
    pdf.ln(5)
    pdf.set_font('Arial', 'B', 12)
    pdf.cell(0, 10, 'TOD-wise Excess Energy Breakdown:', ln=True)
    pdf.set_font('Arial', 'B', 10)
    pdf.cell(30, 8, 'TOD Category', 1)
    pdf.cell(80, 8, 'Description', 1)
    pdf.cell(40, 8, 'Excess Energy (kWh)', 1)
    pdf.ln()
    pdf.set_font('Arial', '', 10)
    
    # Get TOD-wise excess from the dataframe
    tod_excess = pdf_data.groupby('TOD_Category')['Excess'].sum().reset_index()
    
    # Sort by TOD category for consistent display
    tod_excess = tod_excess.sort_values('TOD_Category')
    
    for _, row in tod_excess.iterrows():
        category = row['TOD_Category']
        description = tod_descriptions.get(category, 'Unknown')
        excess = row['Excess']
        pdf.cell(30, 8, category, 1)
        pdf.cell(80, 8, description, 1)
        pdf.cell(40, 8, f"{excess:.4f}", 1)
        pdf.ln()
    
    return tod_excessfrom flask import Flask, render_template, request, send_file
import pandas as pd
from fpdf import FPDF
import io
import os

# Helper function to add TOD summary table with consistent formatting
def add_tod_summary_table(pdf, pdf_data):
    # Define TOD descriptions
    tod_descriptions = {
        'C1': 'Morning Peak',
        'C2': 'Evening Peak',
        'C4': 'Normal Hours',
        'C5': 'Night Hours',
        'Unknown': 'Unknown Time Slot'
    }
    
    # Add TOD-wise excess energy breakdown
    pdf.ln(5)
    pdf.set_font('Arial', 'B', 12)
    pdf.cell(0, 10, 'TOD-wise Excess Energy Breakdown:', ln=True)
    pdf.set_font('Arial', 'B', 10)
    pdf.cell(30, 8, 'TOD Category', 1)
    pdf.cell(80, 8, 'Description', 1)
    pdf.cell(40, 8, 'Excess Energy (kWh)', 1)
    pdf.ln()
    pdf.set_font('Arial', '', 10)
    
    # Get TOD-wise excess from the dataframe
    tod_excess = pdf_data.groupby('TOD_Category')['Excess'].sum().reset_index()
    
    # Sort by TOD category for consistent display
    tod_excess = tod_excess.sort_values('TOD_Category')
    
    for _, row in tod_excess.iterrows():
        category = row['TOD_Category']
        description = tod_descriptions.get(category, 'Unknown')
        excess = row['Excess']
        pdf.cell(30, 8, category, 1)
        pdf.cell(80, 8, description, 1)
        pdf.cell(40, 8, f"{excess:.4f}", 1)
        pdf.ln()
    
    return tod_excessfrom flask import Flask, render_template, request, send_file
import pandas as pd
from fpdf import FPDF
import io
import os

# Helper function to add TOD summary table with consistent formatting
def add_tod_summary_table(pdf, pdf_data):
    # Define TOD descriptions
    tod_descriptions = {
        'C1': 'Morning Peak',
        'C2': 'Evening Peak',
        'C4': 'Normal Hours',
        'C5': 'Night Hours',
        'Unknown': 'Unknown Time Slot'
    }
    
    # Add TOD-wise excess energy breakdown
    pdf.ln(5)
    pdf.set_font('Arial', 'B', 12)
    pdf.cell(0, 10, 'TOD-wise Excess Energy Breakdown:', ln=True)
    pdf.set_font('Arial', 'B', 10)
    pdf.cell(30, 8, 'TOD Category', 1)
    pdf.cell(80, 8, 'Description', 1)
    pdf.cell(40, 8, 'Excess Energy (kWh)', 1)
    pdf.ln()
    pdf.set_font('Arial', '', 10)
    
    # Get TOD-wise excess from the dataframe
    tod_excess = pdf_data.groupby('TOD_Category')['Excess'].sum().reset_index()
    
    # Sort by TOD category for consistent display
    tod_excess = tod_excess.sort_values('TOD_Category')
    
    for _, row in tod_excess.iterrows():
        category = row['TOD_Category']
        description = tod_descriptions.get(category, 'Unknown')
        excess = row['Excess']
        pdf.cell(30, 8, category, 1)
        pdf.cell(80, 8, description, 1)
        pdf.cell(40, 8, f"{excess:.4f}", 1)
        pdf.ln()
    
    return tod_excessfrom flask import Flask, render_template, request, send_file
import pandas as pd
from fpdf import FPDF
import io
import os

# Helper function to add TOD summary table with consistent formatting
def add_tod_summary_table(pdf, pdf_data):
    # Define TOD descriptions
    tod_descriptions = {
        'C1': 'Morning Peak',
        'C2': 'Evening Peak',
        'C4': 'Normal Hours',
        'C5': 'Night Hours',
        'Unknown': 'Unknown Time Slot'
    }
    
    # Add TOD-wise excess energy breakdown
    pdf.ln(5)
    pdf.set_font('Arial', 'B', 12)
    pdf.cell(0, 10, 'TOD-wise Excess Energy Breakdown:', ln=True)
    pdf.set_font('Arial', 'B', 10)
    pdf.cell(30, 8, 'TOD Category', 1)
    pdf.cell(80, 8, 'Description', 1)
    pdf.cell(40, 8, 'Excess Energy (kWh)', 1)
    pdf.ln()
    pdf.set_font('Arial', '', 10)
    
    # Get TOD-wise excess from the dataframe
    tod_excess = pdf_data.groupby('TOD_Category')['Excess'].sum().reset_index()
    
    # Sort by TOD category for consistent display
    tod_excess = tod_excess.sort_values('TOD_Category')
    
    for _, row in tod_excess.iterrows():
        category = row['TOD_Category']
        description = tod_descriptions.get(category, 'Unknown')
        excess = row['Excess']
        pdf.cell(30, 8, category, 1)
        pdf.cell(80, 8, description, 1)
        pdf.cell(40, 8, f"{excess:.4f}", 1)
        pdf.ln()
    
    return tod_excess