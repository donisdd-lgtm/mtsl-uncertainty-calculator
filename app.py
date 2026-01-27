import streamlit as st
import numpy as np
import pandas as pd
import math
from io import BytesIO
from openpyxl import Workbook
from openpyxl.styles import Font, Alignment, PatternFill
from datetime import datetime
from fpdf import FPDF

# Constants
TEMPERATURE_COEFFICIENT = 0.000025
COVERAGE_FACTOR = 2
BMC_FLOOR = 0.083  # Best Measurement Capability floor

st.set_page_config(page_title="MTSL Uncertainty Calc", layout="wide")

st.title("⚡ Uncertainty Calculation Worksheet - MTSL Palakkad")
st.caption("Meter Testing & Standards Laboratory, Palakkad")
st.markdown("---")

# Sidebar for Reference/Certificate Settings
st.sidebar.header("📋 Configuration Parameters")
st.sidebar.markdown("### Reference & Certificate")

ref_standard_accuracy = st.sidebar.number_input(
    "Reference Standard Accuracy (%)",
    value=0.05,
    format="%.4f",
    help="Enter the reference standard accuracy percentage"
)

certificate_uncertainty = st.sidebar.number_input(
    "Certificate Uncertainty (%)",
    value=0.03,
    format="%.4f",
    help="Enter the certificate uncertainty percentage"
)

duc_resolution = st.sidebar.number_input(
    "DUC Resolution",
    value=0.001,
    format="%.4f",
    help="Enter the DUC (Device Under Calibration) resolution"
)

st.sidebar.markdown("### Environmental & Drift Parameters")

temp_difference = st.sidebar.number_input(
    "Temperature Difference (°C)",
    value=0.0,
    format="%.2f",
    help="Enter the temperature difference in degrees Celsius"
)

age_factor = st.sidebar.number_input(
    "Age Factor (%/year)",
    value=0.00001,
    format="%.7f",
    step=0.000001,
    help="Enter the drift percentage per year for the reference standard"
)

years_in_service = st.sidebar.number_input(
    "Years in Service",
    value=10.0,
    step=1.0,
    format="%.1f",
    help="Enter the total years the reference standard has been in service"
)

# Main area for error readings
st.header("📊 Error Readings Input")
st.markdown("Enter the 10 error readings from your measurements:")

# Initialize session state for readings if not exists
if 'readings_df' not in st.session_state:
    st.session_state.readings_df = pd.DataFrame({
        'Reading #': [f"Reading {i}" for i in range(1, 11)],
        'Error Value': [0.0] * 10
    })

# Use data_editor for easy data entry
edited_df = st.data_editor(
    st.session_state.readings_df,
    use_container_width=True,
    hide_index=True,
    num_rows="fixed",
    column_config={
        "Reading #": st.column_config.TextColumn(
            "Reading #",
            disabled=True,
            width="medium"
        ),
        "Error Value": st.column_config.NumberColumn(
            "Error Value",
            format="%.4f",
            width="medium"
        )
    },
    key="error_readings_editor"
)

# Extract error readings from the edited dataframe
error_readings = edited_df['Error Value'].tolist()

st.markdown("---")

# Electrical Parameters Section
st.header("⚡ Electrical Parameters")
st.markdown("Enter the electrical parameters for power calculation:")

col_elec1, col_elec2 = st.columns(2)

with col_elec1:
    ac_type = st.radio(
        "AC Type",
        options=["Single-phase (1φ)", "Three-phase (3φ)"],
        index=0,
        horizontal=True,
        help="Select the AC system type"
    )
    
    voltage = st.number_input(
        "Voltage (V)",
        min_value=0.0,
        value=230.0,
        format="%.2f",
        help="Enter the voltage in volts (V). For three-phase, use line-to-line voltage (V_LL)"
    )
    
with col_elec2:
    current = st.number_input(
        "Current (A)",
        min_value=0.0,
        value=5.0,
        format="%.3f",
        help="Enter the current in amperes (A)"
    )
    
    power_factor = st.number_input(
        "Power Factor",
        min_value=-1.0,
        max_value=1.0,
        value=0.85,
        format="%.3f",
        help="Enter the power factor (-1.0 to +1.0, negative for leading)"
    )

time_hours = st.number_input(
    "Time Duration (hours)",
    min_value=0.0,
    value=1.0,
    format="%.2f",
    help="Enter the time duration in hours for energy calculation"
)

# Calculate Real Power
if ac_type == "Single-phase (1φ)":
    real_power_w = voltage * current * power_factor
    power_formula = "P(W) = V × I × PF"
else:  # Three-phase
    real_power_w = math.sqrt(3) * voltage * current * power_factor
    power_formula = "P(W) = √3 × V_LL × I × PF"

real_power_kw = real_power_w / 1000.0

# Calculate sin(φ) from power factor
# sin(φ) is always positive, but we use the sign of PF to determine reactive power direction
sin_phi = math.copysign(math.sqrt(1 - power_factor**2), power_factor)

# Calculate Reactive Power
if ac_type == "Single-phase (1φ)":
    reactive_power_var = voltage * current * sin_phi
    reactive_formula = "Q(VAr) = V × I × sin(φ)"
else:  # Three-phase
    reactive_power_var = math.sqrt(3) * voltage * current * sin_phi
    reactive_formula = "Q(VAr) = √3 × V_LL × I × sin(φ)"

reactive_power_kvar = reactive_power_var / 1000.0

# Calculate Energy
real_energy_kwh = real_power_kw * time_hours
reactive_energy_kvarh = reactive_power_kvar * time_hours

# Display Power & Energy Calculations
st.subheader("⚡ Power & Energy Calculations")

col_power1, col_power2 = st.columns(2)

with col_power1:
    st.metric(
        "Real Power (kW)",
        f"{real_power_kw:.3f}",
        help=f"Calculated using: {power_formula}"
    )
    
    st.metric(
        "Reactive Power (kVAr)",
        f"{reactive_power_kvar:.3f}",
        help=f"Calculated using: {reactive_formula}"
    )

with col_power2:
    st.metric(
        "Real Energy (kWh)",
        f"{real_energy_kwh:.3f}",
        help=f"Energy = Power × Time = {real_power_kw:.3f} kW × {time_hours:.2f} h"
    )
    
    st.metric(
        "Reactive Energy (kVArh)",
        f"{reactive_energy_kvarh:.3f}",
        help=f"Reactive Energy = Reactive Power × Time = {reactive_power_kvar:.3f} kVAr × {time_hours:.2f} h"
    )

st.markdown("---")

# Calculations
st.header("🔬 Uncertainty Analysis Results")

# Calculate uncertainty components
error_array = np.array(error_readings)

# U1 - Repeatability (Standard Deviation)
U1 = np.std(error_array, ddof=1)  # Using sample standard deviation

# U2 - Reference Standard
U2 = ref_standard_accuracy / math.sqrt(3)

# U3 - Certificate
U3 = certificate_uncertainty / 2

# U4 - Resolution
U4 = duc_resolution / (2 * math.sqrt(3))

# U5 - Temperature Drift
U5 = (TEMPERATURE_COEFFICIENT * temp_difference) / math.sqrt(3)

# U6 - Energy Drift (age-based)
total_drift = age_factor * years_in_service
U6 = total_drift / math.sqrt(3)

# Average Error
average_error = np.mean(error_array)

# Combined Uncertainty (uc)
uc = math.sqrt(U1**2 + U2**2 + U3**2 + U4**2 + U5**2 + U6**2)

# Expanded Uncertainty (U)
expanded_uncertainty = uc * COVERAGE_FACTOR

# Apply BMC floor
final_expanded_uncertainty = max(expanded_uncertainty, BMC_FLOOR)
bmc_applied = expanded_uncertainty < BMC_FLOOR

# Create Uncertainty Budget Table
st.subheader("📋 Uncertainty Budget")

uncertainty_budget = pd.DataFrame({
    "Component": [
        "U1",
        "U2",
        "U3",
        "U4",
        "U5",
        "U6 - Energy Drift (uncertainty included to account for age of the reference standard)"
    ],
    "Description": [
        "Repeatability",
        "Reference Standard Accuracy (uncertainty caused due to the accuracy factor of reference standard used for calibration)",
        "Certificate Uncertainty (uncertainty caused due to the uncertainty reported by the lab where reference standard was calibrated)",
        "Resolution",
        "Temperature Drift (uncertainty caused due to the fact that temperature maintained during the calibration by the lab where reference standard was calibrated is different from the reference temperature)",
        "Energy Drift (uncertainty included to account for age of the reference standard)"
    ],
    "Type": ["A", "B", "B", "B", "B", "B"],
    "Distribution": [
        "Normal",
        "Rectangular",
        "Normal (k=2)",
        "Rectangular",
        "Rectangular",
        "Rectangular"
    ],
    "Value": [
        np.std(error_array, ddof=1),
        ref_standard_accuracy,
        certificate_uncertainty,
        duc_resolution,
        TEMPERATURE_COEFFICIENT * temp_difference,
        total_drift
    ],
    "Divisor": ["1", "√3", "2", "2√3", "√3", "√3"],
    "Standard Uncertainty (ui)": [U1, U2, U3, U4, U5, U6]
})

st.dataframe(
    uncertainty_budget,
    use_container_width=True,
    hide_index=True,
    column_config={
        "Component": st.column_config.TextColumn("Component", width="medium"),
        "Description": st.column_config.TextColumn("Description", width="large"),
        "Type": st.column_config.TextColumn("Type", width="small"),
        "Distribution": st.column_config.TextColumn("Distribution", width="medium"),
        "Value": st.column_config.NumberColumn("Value", format="%.6f", width="medium"),
        "Divisor": st.column_config.TextColumn("Divisor", width="small"),
        "Standard Uncertainty (ui)": st.column_config.NumberColumn("Standard Uncertainty (ui)", format="%.6f", width="medium")
    }
)

st.markdown("---")

# Display results
col_results1, col_results2 = st.columns(2)

with col_results1:
    st.subheader("📈 Uncertainty Components")
    
    st.metric("U1 - Repeatability", f"{U1:.6f}")
    st.metric("U2 - Reference Standard", f"{U2:.6f}")
    st.metric("U3 - Certificate", f"{U3:.6f}")
    st.metric("U4 - Resolution", f"{U4:.6f}")
    st.metric("U5 - Temperature Drift", f"{U5:.6f}")
    st.metric("U6 - Energy Drift (Age)", f"{U6:.6f}")

with col_results2:
    st.subheader("📊 Final Results")
    
    st.metric(
        "Average Error", 
        f"{average_error:.6f}",
        help="Mean of the 10 error readings"
    )
    
    st.metric(
        "Combined Uncertainty (uc)", 
        f"{uc:.6f}",
        help="Root sum of squares of all uncertainty components"
    )
    
    st.metric(
        "Expanded Uncertainty (U)", 
        f"{final_expanded_uncertainty:.6f}",
        delta=f"k = {COVERAGE_FACTOR}" + (" | BMC Floor Applied" if bmc_applied else ""),
        help=f"Combined uncertainty multiplied by coverage factor (k={COVERAGE_FACTOR}). BMC floor = {BMC_FLOOR}%"
    )

# Detailed breakdown
st.markdown("---")
st.subheader("📋 Detailed Calculation Breakdown")

with st.expander("View Calculation Details"):
    st.markdown("### Input Parameters")
    st.write(f"- **Error Readings:** {error_readings}")
    st.write(f"- **Reference Standard Accuracy:** {ref_standard_accuracy}%")
    st.write(f"- **Certificate Uncertainty:** {certificate_uncertainty}%")
    st.write(f"- **DUC Resolution:** {duc_resolution}")
    st.write(f"- **Temperature Difference:** {temp_difference}°C")
    st.write(f"- **Age Factor:** {age_factor}%/year")
    st.write(f"- **Years in Service:** {years_in_service} years")
    st.write(f"- **AC Type:** {ac_type}")
    st.write(f"- **Voltage:** {voltage} V")
    st.write(f"- **Current:** {current} A")
    st.write(f"- **Power Factor:** {power_factor}")
    st.write(f"- **Time Duration:** {time_hours} hours")
    st.write(f"- **Real Power:** {real_power_kw:.3f} kW (calculated using {power_formula})")
    st.write(f"- **Reactive Power:** {reactive_power_kvar:.3f} kVAr (calculated using {reactive_formula})")
    st.write(f"- **Real Energy:** {real_energy_kwh:.3f} kWh")
    st.write(f"- **Reactive Energy:** {reactive_energy_kvarh:.3f} kVArh")
    
    st.markdown("### Formulas Used")
    st.latex(r"U_1 = \\sigma_{readings}")
    st.latex(r"U_2 = \\frac{Reference\;Accuracy}{\\sqrt{3}}")
    st.latex(r"U_3 = \\frac{Certificate\;Uncertainty}{2}")
    st.latex(r"U_4 = \\frac{DUC\;Resolution}{2\\sqrt{3}}")
    st.latex(r"U_5 = \\frac{0.000025 \times Temp\;Difference}{\\sqrt{3}}")
    st.latex(r"U_6 = \\frac{Energy\;Drift}{\\sqrt{3}}")
    st.latex(r"u_c = \\sqrt{U_1^2 + U_2^2 + U_3^2 + U_4^2 + U_5^2 + U_6^2}")
    st.latex(r"U = u_c \times k \quad (k=2)")
    
    st.markdown("### Component Contributions")
    contributions = {
        "U1 (Repeatability)": U1**2,
        "U2 (Ref Standard)": U2**2,
        "U3 (Certificate)": U3**2,
        "U4 (Resolution)": U4**2,
        "U5 (Temp Drift)": U5**2,
        "U6 (Energy Drift - Age)": U6**2
    }
    
    total_variance = sum(contributions.values())
    
    for component, variance in contributions.items():
        if total_variance > 0:
            percentage = (variance / total_variance) * 100
            st.write(f"- **{component}:** {variance:.8f} ({percentage:.2f}% contribution)")
        else:
            st.write(f"- **{component}:** {variance:.8f}")

# Excel Export Function
def create_excel_report():
    output = BytesIO()
    wb = Workbook()
    ws = wb.active
    ws.title = "Uncertainty Calculation"
    
    # Define styles
    header_fill = PatternFill(start_color="0068C9", end_color="0068C9", fill_type="solid")
    header_font = Font(bold=True, color="FFFFFF", size=12)
    title_font = Font(bold=True, size=14)
    bold_font = Font(bold=True)
    center_alignment = Alignment(horizontal="center", vertical="center")
    
    row = 1
    
    # Title
    ws.merge_cells(f'A{row}:E{row}')
    cell = ws[f'A{row}']
    cell.value = "Uncertainty Calculation Worksheet - Electrical Lab"
    cell.font = title_font
    cell.alignment = center_alignment
    row += 1
    
    ws[f'A{row}'] = f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"
    row += 2
    
    # Input Parameters Section
    ws[f'A{row}'] = "INPUT PARAMETERS"
    ws[f'A{row}'].font = bold_font
    row += 1
    
    ws[f'A{row}'] = "Reference Standard Accuracy (%)"
    ws[f'B{row}'] = ref_standard_accuracy
    row += 1
    
    ws[f'A{row}'] = "Certificate Uncertainty (%)"
    ws[f'B{row}'] = certificate_uncertainty
    row += 1
    
    ws[f'A{row}'] = "DUC Resolution"
    ws[f'B{row}'] = duc_resolution
    row += 1
    
    ws[f'A{row}'] = "Temperature Difference (°C)"
    ws[f'B{row}'] = temp_difference
    row += 1
    
    ws[f'A{row}'] = "Age Factor (%/year)"
    ws[f'B{row}'] = age_factor
    row += 1
    
    ws[f'A{row}'] = "Years in Service"
    ws[f'B{row}'] = years_in_service
    row += 1
    
    ws[f'A{row}'] = "AC Type"
    ws[f'B{row}'] = ac_type
    row += 1
    
    ws[f'A{row}'] = "Voltage (V)"
    ws[f'B{row}'] = voltage
    row += 1
    
    ws[f'A{row}'] = "Current (A)"
    ws[f'B{row}'] = current
    row += 1
    
    ws[f'A{row}'] = "Power Factor"
    ws[f'B{row}'] = power_factor
    row += 1
    
    ws[f'A{row}'] = "Time Duration (hours)"
    ws[f'B{row}'] = time_hours
    row += 1
    
    ws[f'A{row}'] = "Real Power (kW)"
    ws[f'B{row}'] = real_power_kw
    row += 1
    
    ws[f'A{row}'] = "Reactive Power (kVAr)"
    ws[f'B{row}'] = reactive_power_kvar
    row += 1
    
    ws[f'A{row}'] = "Real Energy (kWh)"
    ws[f'B{row}'] = real_energy_kwh
    row += 1
    
    ws[f'A{row}'] = "Reactive Energy (kVArh)"
    ws[f'B{row}'] = reactive_energy_kvarh
    row += 2
    
    # Error Readings Section
    ws[f'A{row}'] = "ERROR READINGS"
    ws[f'A{row}'].font = bold_font
    row += 1
    
    ws[f'A{row}'] = "Reading #"
    ws[f'B{row}'] = "Error Value"
    ws[f'A{row}'].font = bold_font
    ws[f'B{row}'].font = bold_font
    row += 1
    
    for i, reading in enumerate(error_readings, 1):
        ws[f'A{row}'] = f"Reading {i}"
        ws[f'B{row}'] = reading
        row += 1
    
    row += 1
    
    # Uncertainty Budget Table
    ws[f'A{row}'] = "UNCERTAINTY BUDGET"
    ws[f'A{row}'].font = bold_font
    row += 1
    
    # Headers
    headers = ["Component", "Description", "Type", "Distribution", "Value", "Divisor", "Standard Uncertainty (ui)"]
    for col_idx, header in enumerate(headers, 1):
        cell = ws.cell(row=row, column=col_idx)
        cell.value = header
        cell.font = header_font
        cell.fill = header_fill
        cell.alignment = center_alignment
    row += 1
    
    # Data rows
    for _, budget_row in uncertainty_budget.iterrows():
        ws[f'A{row}'] = budget_row['Component']
        ws[f'B{row}'] = budget_row['Description']
        ws[f'C{row}'] = budget_row['Type']
        ws[f'D{row}'] = budget_row['Distribution']
        ws[f'E{row}'] = budget_row['Value']
        ws[f'F{row}'] = budget_row['Divisor']
        ws[f'G{row}'] = budget_row['Standard Uncertainty (ui)']
        row += 1
    
    row += 1
    
    # Final Results Section
    ws[f'A{row}'] = "FINAL RESULTS"
    ws[f'A{row}'].font = bold_font
    row += 1
    
    ws[f'A{row}'] = "Average Error (%)"
    ws[f'B{row}'] = average_error
    row += 1
    
    ws[f'A{row}'] = "Combined Uncertainty - uc (%)"
    ws[f'B{row}'] = uc
    row += 1
    
    ws[f'A{row}'] = "Expanded Uncertainty - U (k={COVERAGE_FACTOR}) (%)"
    ws[f'B{row}'] = final_expanded_uncertainty
    row += 1
    
    if bmc_applied:
        ws[f'A{row}'] = "BMC Floor Applied"
        ws[f'B{row}'] = f"Calculated: {expanded_uncertainty:.6f}% | Final: {BMC_FLOOR}%"
        ws[f'A{row}'].font = Font(italic=True, color="FF6600")
        row += 1
    
    row += 1
    
    # Final Result Statement
    ws.merge_cells(f'A{row}:E{row}')
    cell = ws[f'A{row}']
    cell.value = f"Result: {average_error:.4f}% ± {final_expanded_uncertainty:.4f}% (k={COVERAGE_FACTOR})"
    cell.font = Font(bold=True, size=14, color="0068C9")
    cell.alignment = center_alignment
    
    # Adjust column widths
    ws.column_dimensions['A'].width = 30
    ws.column_dimensions['B'].width = 50
    ws.column_dimensions['C'].width = 8
    ws.column_dimensions['D'].width = 18
    ws.column_dimensions['E'].width = 15
    ws.column_dimensions['F'].width = 10
    ws.column_dimensions['G'].width = 25
    
    wb.save(output)
    output.seek(0)
    return output

# PDF Export Function
def create_pdf_report():
    pdf = FPDF()
    pdf.add_page()
    pdf.set_auto_page_break(auto=True, margin=15)
    
    # Helper function to clean special characters
    def clean_text(text):
        text = str(text)
        text = text.replace("±", "+/-")
        text = text.replace("√", "sqrt")
        text = text.replace("°", " deg")
        text = text.replace("₁", "1")
        text = text.replace("₂", "2")
        text = text.replace("₃", "3")
        text = text.replace("₄", "4")
        text = text.replace("₅", "5")
        text = text.replace("₆", "6")
        text = text.replace("φ", "phi")
        text = text.replace("×", "x")
        return text
    
    # Header
    pdf.set_font("Arial", "B", 16)
    pdf.cell(0, 10, "Uncertainty Calculation Worksheet - MTSL Palakkad", ln=True, align="C")
    pdf.set_font("Arial", "I", 10)
    pdf.cell(0, 6, "Meter Testing & Standards Laboratory, Palakkad", ln=True, align="C")
    pdf.set_font("Arial", "", 9)
    pdf.cell(0, 6, f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}", ln=True, align="C")
    pdf.ln(5)
    
    # Section 1 - Input Parameters
    pdf.set_font("Arial", "B", 12)
    pdf.cell(0, 8, "INPUT PARAMETERS", ln=True)
    pdf.set_font("Arial", "", 10)
    
    pdf.cell(0, 6, f"Reference Standard Accuracy: {ref_standard_accuracy:.4f}%", ln=True)
    pdf.cell(0, 6, f"Certificate Uncertainty: {certificate_uncertainty:.4f}%", ln=True)
    pdf.cell(0, 6, f"DUC Resolution: {duc_resolution:.4f}", ln=True)
    pdf.cell(0, 6, f"Temperature Difference: {temp_difference:.2f} deg C", ln=True)
    pdf.cell(0, 6, f"Age Factor: {age_factor:.5f}%/year", ln=True)
    pdf.cell(0, 6, f"Years in Service: {years_in_service:.1f} years", ln=True)
    pdf.cell(0, 6, f"AC Type: {clean_text(ac_type)}", ln=True)
    pdf.cell(0, 6, f"Voltage: {voltage:.2f} V", ln=True)
    pdf.cell(0, 6, f"Current: {current:.3f} A", ln=True)
    pdf.cell(0, 6, f"Power Factor: {power_factor:.3f}", ln=True)
    pdf.cell(0, 6, f"Time Duration: {time_hours:.2f} hours", ln=True)
    pdf.cell(0, 6, f"Real Power: {real_power_kw:.3f} kW (Formula: {clean_text(power_formula)})", ln=True)
    pdf.cell(0, 6, f"Reactive Power: {reactive_power_kvar:.3f} kVAr (Formula: {clean_text(reactive_formula)})", ln=True)
    pdf.cell(0, 6, f"Real Energy: {real_energy_kwh:.3f} kWh", ln=True)
    pdf.cell(0, 6, f"Reactive Energy: {reactive_energy_kvarh:.3f} kVArh", ln=True)
    pdf.ln(3)
    
    # Error Readings
    pdf.set_font("Arial", "B", 11)
    pdf.cell(0, 6, "Error Readings (10 measurements):", ln=True)
    pdf.set_font("Arial", "", 9)
    
    # Display readings in 2 rows
    for i in range(0, 10, 5):
        readings_line = "  ".join([f"R{j+1}: {error_readings[j]:.4f}" for j in range(i, min(i+5, 10))])
        pdf.cell(0, 5, readings_line, ln=True)
    pdf.ln(3)
    
    # Section 2 - Uncertainty Budget
    pdf.set_font("Arial", "B", 12)
    pdf.cell(0, 8, "UNCERTAINTY BUDGET", ln=True)
    pdf.ln(2)
    
    # Table headers
    pdf.set_font("Arial", "B", 9)
    pdf.set_fill_color(0, 104, 201)
    pdf.set_text_color(255, 255, 255)
    pdf.cell(18, 7, "Comp.", border=1, fill=True, align="C")
    pdf.cell(20, 7, "Type", border=1, fill=True, align="C")
    pdf.cell(32, 7, "Distribution", border=1, fill=True, align="C")
    pdf.cell(20, 7, "Value", border=1, fill=True, align="C")
    pdf.cell(18, 7, "Divisor", border=1, fill=True, align="C")
    pdf.cell(25, 7, "Std Unc (ui)", border=1, fill=True, align="C")
    pdf.ln()
    
    pdf.set_text_color(0, 0, 0)
    pdf.set_font("Arial", "", 8)
    
    # Table data
    for idx, row in uncertainty_budget.iterrows():
        comp_name = clean_text(row['Component'])
        # Shorten component name for table
        if "U1" in comp_name:
            comp_short = "U1"
        elif "U2" in comp_name:
            comp_short = "U2"
        elif "U3" in comp_name:
            comp_short = "U3"
        elif "U4" in comp_name:
            comp_short = "U4"
        elif "U5" in comp_name:
            comp_short = "U5"
        elif "U6" in comp_name:
            comp_short = "U6"
        else:
            comp_short = comp_name[:10]
        
        pdf.cell(18, 6, comp_short, border=1, align="C")
        pdf.cell(20, 6, clean_text(row['Type']), border=1, align="C")
        pdf.cell(32, 6, clean_text(row['Distribution']), border=1, align="C")
        pdf.cell(20, 6, f"{row['Value']:.6f}", border=1, align="C")
        pdf.cell(18, 6, clean_text(row['Divisor']), border=1, align="C")
        pdf.cell(25, 6, f"{row['Standard Uncertainty (ui)']:.6f}", border=1, align="C")
        pdf.ln()
    
    pdf.ln(3)
    
    # Component Descriptions
    pdf.set_font("Arial", "B", 11)
    pdf.cell(0, 6, "Component Descriptions:", ln=True)
    pdf.set_font("Arial", "", 8)
    
    for idx, row in uncertainty_budget.iterrows():
        comp_name = clean_text(row['Component'])
        description = clean_text(row['Description'])
        pdf.multi_cell(0, 4, f"{comp_name}: {description}")
        pdf.ln(1)
    
    pdf.ln(2)
    
    # Section 3 - Calculation Formulas
    pdf.set_font("Arial", "B", 12)
    pdf.cell(0, 8, "CALCULATION FORMULAS", ln=True)
    pdf.set_font("Arial", "", 9)
    
    pdf.cell(0, 5, f"U1 (Repeatability) = Standard Deviation of readings = {U1:.6f}", ln=True)
    pdf.cell(0, 5, f"U2 (Ref Standard) = Ref Accuracy / sqrt(3) = {ref_standard_accuracy:.6f} / 1.732 = {U2:.6f}", ln=True)
    pdf.cell(0, 5, f"U3 (Certificate) = Cert Uncertainty / 2 = {certificate_uncertainty:.6f} / 2 = {U3:.6f}", ln=True)
    pdf.cell(0, 5, f"U4 (Resolution) = DUC Res / (2 * sqrt(3)) = {duc_resolution:.6f} / 3.464 = {U4:.6f}", ln=True)
    temp_val = TEMPERATURE_COEFFICIENT * temp_difference
    pdf.cell(0, 5, f"U5 (Temp Drift) = (0.000025 * {temp_difference:.2f}) / sqrt(3) = {temp_val:.8f} / 1.732 = {U5:.8f}", ln=True)
    pdf.cell(0, 5, f"U6 (Energy Drift) = (Age Factor * Years) / sqrt(3) = ({age_factor:.5f} * {years_in_service:.0f}) / 1.732 = {U6:.8f}", ln=True)
    pdf.ln(3)
    
    pdf.cell(0, 5, f"Combined Uncertainty (uc) = sqrt(U1^2 + U2^2 + ... + U6^2) = {uc:.6f}%", ln=True)
    pdf.cell(0, 5, f"Expanded Uncertainty (U) = k * uc = {COVERAGE_FACTOR} * {uc:.6f} = {expanded_uncertainty:.6f}%", ln=True)
    
    if bmc_applied:
        pdf.ln(2)
        pdf.set_font("Arial", "B", 9)
        pdf.set_text_color(255, 102, 0)
        pdf.cell(0, 5, f"BMC Floor Applied: Calculated U ({expanded_uncertainty:.4f}%) set to {BMC_FLOOR:.3f}%.", ln=True)
        pdf.set_text_color(0, 0, 0)
    
    pdf.ln(5)
    
    # Section 4 - Final Result
    pdf.set_font("Arial", "B", 14)
    pdf.set_text_color(0, 104, 201)
    pdf.cell(0, 10, f"FINAL RESULT: {average_error:.4f}% +/- {final_expanded_uncertainty:.4f}% (k={COVERAGE_FACTOR})", ln=True, align="C", border=1)
    pdf.set_text_color(0, 0, 0)
    
    # Footer
    pdf.ln(5)
    pdf.set_font("Arial", "I", 8)
    pdf.cell(0, 5, "Coverage Factor k=2 (95.45% confidence level)", ln=True, align="C")
    
    # Output to BytesIO
    output = BytesIO()
    pdf_output = pdf.output(dest='S').encode('latin-1')
    output.write(pdf_output)
    output.seek(0)
    return output

# Download Button
st.markdown("---")
col_download1, col_download2 = st.columns(2)

with col_download1:
    excel_file = create_excel_report()
    st.download_button(
        label="📥 Download Calculation as Excel",
        data=excel_file,
        file_name=f"uncertainty_calculation_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )

with col_download2:
    pdf_file = create_pdf_report()
    st.download_button(
        label="📄 Download PDF Report",
        data=pdf_file,
        file_name=f"MTSL_Uncertainty_Report_{datetime.now().strftime('%Y%m%d_%H%M%S')}.pdf",
        mime="application/pdf",
        use_container_width=True
    )

# Final Result Display
st.markdown("---")
st.markdown(
    f"""
    <div style='text-align: center; padding: 20px; background-color: #f0f2f6; border-radius: 10px; margin: 20px 0;'>
        <h2 style='color: #0068c9; margin: 0;'>
            <strong>Result: {average_error:.4f}% ± {final_expanded_uncertainty:.4f}% (k={COVERAGE_FACTOR})</strong>
        </h2>
    </div>
    """, 
    unsafe_allow_html=True
)

# Detailed Step-by-Step Calculation
with st.expander("🔍 Show Detailed Step-by-Step Calculation"):
    st.markdown("### Mathematical Proof and Derivation")
    st.markdown("---")
    
    # Average Error
    st.markdown("#### **Average Error**")
    sum_of_errors = sum(error_readings)
    st.latex(r"\text{Average Error} = \frac{\sum_{i=1}^{10} \text{Error}_i}{10}")
    st.latex(f"= \\frac{{{sum_of_errors:.6f}}}{{10}} = {average_error:.6f} \%")
    st.markdown("")
    
    # U1 - Repeatability
    st.markdown("#### **U₁ - Repeatability (Standard Deviation)**")
    st.latex(r"U_1 = \\sigma = \\sqrt{\\frac{\sum_{i=1}^{n}(x_i - \bar{x})^2}{n-1}}")
    st.markdown(f"Where readings are: '{{[f'{{r:.4f}}' for r in error_readings]}}'")
    st.latex(f"U_1 = {U1:.6f}")
    st.markdown("")
    
    # U2 - Reference Standard
    st.markdown("#### **U₂ - Reference Standard Accuracy (uncertainty caused due to the accuracy factor of reference standard used for calibration)**")
    st.latex(r"U_{2} = \\frac{Reference Accuracy}{\\sqrt{3}}")
    st.latex(f"= \\frac{{{ref_standard_accuracy:.6f}}}{{\\sqrt{{3}}}} = \\frac{{{ref_standard_accuracy:.6f}}}{{1.732051}} = {U2:.6f} \%")
    st.markdown("")
    
    # U3 - Certificate
    st.markdown("#### **U₃ - Certificate Uncertainty (uncertainty caused due to the uncertainty reported by the lab where reference standard was calibrated)**")
    st.latex(r"U_{3} = \\frac{Certificate Uncertainty}{2}")
    st.latex(f"= \\frac{{{certificate_uncertainty:.6f}}}{{2}} = {U3:.6f} \%")
    st.markdown("")
    
    # U4 - Resolution
    st.markdown("#### **U₄ - DUC Resolution**")
    st.latex(r"U_{4} = \\frac{DUC Resolution}{2\\sqrt{3}}")
    st.latex(f"= \\frac{{{duc_resolution:.6f}}}{{2 \\times 1.732051}} = \\frac{{{duc_resolution:.6f}}}{{3.464102}} = {U4:.6f}")
    st.markdown("")
    
    # U5 - Temperature Drift
    st.markdown("#### **U₅ - Temperature Drift (uncertainty caused due to the fact that temperature maintained during the calibration by the lab where reference standard was calibrated is different from the reference temperature)**")
    st.latex(r"U_{5} = \\frac{Temp Coeff \times \Delta T}{\\sqrt{3}}")
    temp_value = TEMPERATURE_COEFFICIENT * temp_difference
    st.latex(f"= \\frac{{0.000025 \times {temp_difference:.2f}}}{{\\sqrt{{3}}}} = \\frac{{{temp_value:.8f}}}{{1.732051}} = {U5:.8f}")
    st.markdown("")
    
    # U6 - Energy Drift (Age-Based)
    st.markdown("#### **U₆ - Energy Drift (uncertainty included to account for age of the reference standard)**")
    st.latex(r"U_{drift} = \\frac{Age Factor \times Years}{\\sqrt{3}}")
    st.latex(f"= \\frac{{{age_factor:.5f} \times {years_in_service:.0f}}}{{1.732051}} = {U6:.8f} \%")
    st.markdown("")
    
    st.markdown("---")
    
    # Combined Uncertainty
    st.markdown("#### **Combined Uncertainty (uᴄ)**")
    st.latex(r"u_c = \\sqrt{U_1^2 + U_2^2 + U_3^2 + U_4^2 + U_5^2 + U_6^2}")
    st.latex(f"= \\sqrt{{{U1:.6f}^2 + {U2:.6f}^2 + {U3:.6f}^2 + {U4:.6f}^2 + {U5:.8f}^2 + {U6:.8f}^2}}")
    sum_of_squares = U1**2 + U2**2 + U3**2 + U4**2 + U5**2 + U6**2
    st.latex(f"= \\sqrt{{{sum_of_squares:.12f}}} = {uc:.6f} \%")
    st.markdown("")
    
    # Expanded Uncertainty
    st.markdown("#### **Expanded Uncertainty (U)**")
    st.latex(r"U = k \times u_c")
    st.latex(f"= {COVERAGE_FACTOR} \times {uc:.6f} = {expanded_uncertainty:.6f} \%")
    
    # BMC Floor Check
    if bmc_applied:
        st.markdown("#### **⚠️ Best Measurement Capability (BMC) Floor Applied**")
        st.warning(f"**Note:** Calculated U ({expanded_uncertainty:.4f}%) is below the BMC limit. Final Result set to {BMC_FLOOR:.3f}%.")
        st.latex(f"U_{{final}} = \max(U_{{calculated}}, \text{{BMC Floor}}) = \max({expanded_uncertainty:.6f}, {BMC_FLOOR}) = {final_expanded_uncertainty:.4f} \%")
        st.markdown("")
    
    st.markdown("---")
    st.markdown("### **Final Result**")
    st.latex(f"\text{{Result}} = {average_error:.4f}\% \pm {final_expanded_uncertainty:.4f}\% \quad (k={COVERAGE_FACTOR})")


# Footer
st.markdown("---")
st.markdown(
    """
    <div style='text-align: center; color: gray; font-size: 0.9em;'>
    Uncertainty Calculation Worksheet | Electrical Lab | Coverage Factor k=2
    </div>
    """, 
    unsafe_allow_html=True
)