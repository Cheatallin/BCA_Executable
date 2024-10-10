import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import statsmodels.api as sm
import tkinter as tk
from tkinter import filedialog
from fpdf import FPDF

# Function to open file dialog and read the selected file
def load_file():
    root = tk.Tk()
    root.withdraw()  # Hide the root window
    file_path = filedialog.askopenfilename(
        filetypes=[("Excel files", "*.xlsx"), ("CSV files", "*.csv")],
        title="Select a file"
    )
    if file_path.endswith('.xlsx'):
        return pd.read_excel(file_path)
    elif file_path.endswith('.csv'):
        return pd.read_csv(file_path)
    else:
        raise ValueError("Unsupported file type")

# Load the data
BCA1 = load_file()

# Calculating each readings means
Blank = BCA1['Blank'].mean()
C005 = BCA1['0.05g/l'].mean()
C01 = BCA1['0.1g/l'].mean()
C02 = BCA1['0.2g/l'].mean()
C04 = BCA1['0.4g/l'].mean()

# Identify all "Sample" columns
sample_columns = [col for col in BCA1.columns if col.startswith('Sample')]
sample_means = {col: BCA1[col].mean() for col in sample_columns}

# Deplete blank from each readings
C005_mod = C005 - Blank
C01_mod = C01 - Blank
C02_mod = C02 - Blank
C04_mod = C04 - Blank
sample_mods = {col: mean - Blank for col, mean in sample_means.items()}

# Creating data frame for standard line
BCA_Line = pd.DataFrame({
    "Abs": [C005_mod, C01_mod, C02_mod, C04_mod],
    "Conc": [0.05, 0.1, 0.2, 0.4]
})

# Export logbook to a text file
with open("logbook.txt", "w") as file:
    file.write("Abs Conc\n")
    for index, row in BCA_Line.iterrows():
        file.write(f"{index + 1} {row['Abs']:.8f} {row['Conc']}\n")

# Creating Standard Curve
X = BCA_Line['Abs']
X = sm.add_constant(X)  # Adds a constant term to the predictor
Y = BCA_Line['Conc']
Std_Line = sm.OLS(Y, X).fit()

# Append linear model summary to the text file
with open("logbook.txt", "a") as file:
    file.write("\nCall:\nlm(formula = BCA_Line$Conc ~ BCA_Line$Abs)\n")
    file.write("Coefficients:\n")
    file.write(str(Std_Line.params) + "\n")
    file.write("\n" + str(Std_Line.summary()) + "\n")

# Calculate concentrations for all samples
sample_concentrations = {col: mod * Std_Line.params['Abs'] + Std_Line.params['const'] for col, mod in sample_mods.items()}

# Append sample concentrations to the text file
with open("logbook.txt", "a") as file:
    for col, conc in sample_concentrations.items():
        file.write(f"\n{col}: {conc}\n")

# Create a PDF class
class PDF(FPDF):
    def header(self):
        self.set_font('Arial', 'B', 12)
        self.cell(0, 10, 'Logbook', 0, 1, 'C')

    def footer(self):
        self.set_y(-15)
        self.set_font('Arial', 'I', 8)
        self.cell(0, 10, f'Page {self.page_no()}', 0, 0, 'C')

# Create a PDF instance
pdf = PDF()
pdf.add_page()
pdf.set_font("Arial", size=12)

# Add logbook content to PDF
pdf.cell(0, 10, "Abs Conc", ln=True)
for index, row in BCA_Line.iterrows():
    pdf.cell(0, 10, f"{index + 1} {row['Abs']:.8f} {row['Conc']}", ln=True)

pdf.cell(0, 10, "\nCall:\nlm(formula = BCA_Line$Conc ~ BCA_Line$Abs)\n", ln=True)
pdf.cell(0, 10, "Coefficients:", ln=True)
pdf.cell(0, 10, str(Std_Line.params), ln=True)
pdf.multi_cell(0, 10, str(Std_Line.summary()))

for col, conc in sample_concentrations.items():
    pdf.cell(0, 10, f"\n{col}: {conc}", ln=True)

# Save the plot as an image
plt.scatter(BCA_Line['Conc'], BCA_Line['Abs'], s=20)
plt.plot(BCA_Line['Conc'], Std_Line.predict(X), color='red')
plt.title("BSA Standard Line")
plt.axvline(x=0.05, color='grey', linestyle='--')
plt.axhline(y=0, color='grey', linestyle='--')
plt.xlabel("BSA Concentration (g/l)")
plt.ylabel("Absorbance")
plt.savefig("BCA_Graph.jpg")
plt.close()

# Add the graph to the PDF
pdf.add_page()
pdf.image("BCA_Graph.jpg", x=10, y=10, w=190)

# Save the PDF
pdf.output("logbook.pdf")