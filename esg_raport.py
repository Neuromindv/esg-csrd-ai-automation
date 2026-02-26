import pandas as pd
from openpyxl import Workbook, load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from openpyxl.styles import PatternFill
from openpyxl.chart import BarChart, Reference
import os

def create_dummy_data(filename="dane_wejsciowe.xlsx"):
    """Generuje plik wejściowy, aby skrypt działał 'out of the box'."""
    data = {
        'Company': ['Firma Alpha', 'Firma Beta', 'Firma Gamma', 'Firma Delta'],
        'Revenue': [150.0, 320.0, 80.0, 500.0], # W milionach EUR
        'Emissions_Scope1': [4500, 25000, 2000, 15000], # Tony CO2e
        'Emissions_Scope2': [1500, 8000, 500, 5000],    # Tony CO2e
        'Employees': [120, 850, 45, 1200],
        'Sector': ['IT', 'Manufacturing', 'Services', 'Logistics']
    }
    df = pd.DataFrame(data)
    df.to_excel(filename, index=False)
    print(f"Utworzono plik wejściowy: {filename}")

def generate_esg_report(input_file, output_file):
    # 1. Wczytanie danych z wykorzystaniem Pandas
    df = pd.read_excel(input_file)
    
    # 2. Obliczenia
    # Intensywność węglowa: (Scope 1 + Scope 2) / Przychód (tCO2e / mln EUR) - Zgodnie z ESRS E1-6
    df['Total_Emissions'] = df['Emissions_Scope1'] + df['Emissions_Scope2']
    df['Carbon_Intensity'] = round(df['Total_Emissions'] / df['Revenue'], 2)
    
    # Intensywność zatrudnienia: Pracownicy / Przychód (liczba pracowników na 1 mln EUR przychodu)
    df['Employee_Intensity'] = round(df['Employees'] / df['Revenue'], 2)
    
    # Przebudowa kolejności kolumn dla lepszej czytelności
    cols = ['Company', 'Sector', 'Revenue', 'Employees', 'Emissions_Scope1', 
            'Emissions_Scope2', 'Total_Emissions', 'Carbon_Intensity', 'Employee_Intensity']
    df = df[cols]

    # 3. Zapis i formatowanie z użyciem openpyxl
    wb = Workbook()
    
    # --- Zakładka z danymi ---
    ws_data = wb.active
    ws_data.title = "Dane ESG (ESRS)"
    
    # Wpisanie danych z Pandas do Openpyxl
    for r in dataframe_to_rows(df, index=False, header=True):
        ws_data.append(r)
        
    # Ustawienie szerokości kolumn
    for col in ws_data.columns:
        max_length = 0
        column = col[0].column_letter # Zwraca literę kolumny
        for cell in col:
            try:
                if len(str(cell.value)) > max_length:
                    max_length = len(str(cell.value))
            except:
                pass
        adjusted_width = (max_length + 2)
        ws_data.column_dimensions[column].width = adjusted_width

    # Definicja kolorów (Color Coding)
    green_fill = PatternFill(start_color="C6EFCE", end_color="C6EFCE", fill_type="solid")
    yellow_fill = PatternFill(start_color="FFEB9C", end_color="FFEB9C", fill_type="solid")
    red_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")

    # Aplikacja Color Coding dla kolumny Carbon_Intensity (Kolumna H to 8 kolumna)
    # UWAGA: Progi są czysto przykładowe
    for row in range(2, ws_data.max_row + 1):
        cell = ws_data.cell(row=row, column=8)
        val = cell.value
        if isinstance(val, (int, float)):
            if val < 50:
                cell.fill = green_fill
            elif 50 <= val <= 80:
                cell.fill = yellow_fill
            else:
                cell.fill = red_fill

    # --- Zakładka z wykresem ---
    ws_chart = wb.create_sheet(title="Dashboard")
    
    chart = BarChart()
    chart.title = "Carbon Intensity według spółek (tCO2e / mln EUR)"
    chart.y_axis.title = 'Intensywność Węglowa'
    chart.x_axis.title = 'Spółka'
    
    # Dane dla wykresu (Kolumna H - Carbon Intensity)
    data_ref = Reference(ws_data, min_col=8, min_row=1, max_row=ws_data.max_row)
    # Kategorie osi X (Kolumna A - Company)
    cats_ref = Reference(ws_data, min_col=1, min_row=2, max_row=ws_data.max_row)
    
    chart.add_data(data_ref, titles_from_data=True)
    chart.set_categories(cats_ref)
    chart.width = 15
    chart.height = 8
    
    ws_chart.add_chart(chart, "B2")

    # Zapis pliku
    wb.save(output_file)
    print(f"Raport z sukcesem zapisany jako: {output_file}")

if __name__ == "__main__":
    input_filename = "dane_wejsciowe.xlsx"
    output_filename = "Raport_ESG_CSRD.xlsx"
    
    create_dummy_data(input_filename)
    generate_esg_report(input_filename, output_filename)
