import pandas as pd
import os 
from pathlib import Path 

root_dir = Path(".").resolve().parent
filename = "PRUEBA ANALISTA PYTHON.xlsx"

def get_data(filename):
    file_path = os.path.join(root_dir, "Data", "Raw", filename)
    Data_inicial = pd.read_excel(file_path,sheet_name="FORMATO INICIAL")
    Data_agentes = pd.read_excel(file_path,sheet_name="BASE AGENTES")
    return Data_inicial,Data_agentes

def generate_report(Data_inicial,Data_agentes):
    Data_inicial['Individual Name'] = Data_inicial['Individual Name'].str.split()
    Data_inicial['individual:first name'] = Data_inicial['Individual Name'].apply(lambda x: x[-1] if len(x) == 3 or len(x) == 2 else x[-2] if len(x) == 4 or len == 5 or len == 6 else None)
    Data_inicial['Individual Name'] = Data_inicial['Individual Name'].apply(lambda x: ' '.join(x))
    Data_inicial['Individual Name'] = Data_inicial.apply(lambda row: row['Individual Name'].replace(row['individual:first name'], '') if pd.notna(row['Individual Name']) and pd.notna(row['individual:first name']) else row['Individual Name'], axis=1)
    Data_inicial['Agent: First Name'] = Data_inicial['Agent Name'].str.split().str[0]
    Data_inicial['Agent: Last Name'] = Data_inicial['Agent Name'].str.split().str[-1]
    # Combinar DataFrames usando 'Agent:Name' como clave
    result = pd.merge(Data_inicial, Data_agentes[['Agent:Name', 'Agent: External ID', 'Agency: Agency Name']],
                  left_on='Agent #', right_on='Agent: External ID', how='left')
    df_subset = result[['individual:first name', 'Individual Name', 'Carrier: Carrier Name', 
                'Policy: Coverage Type', 'Policy: Policy Number', 'Policy: Status', 
                'Policy: Effective Date', 'Agent #','Agent: First Name', 'Agent: Last Name', 
                'Agency: Agency Name']]
    df_subset = df_subset.rename(columns={'Individual Name': 'individual:last name', 'Agent #': 'Agent: External ID'})
    return df_subset

def save_date(df_subset,filename):
    out_name = "reporte.csv"  
    out_path = os.path.join(root_dir,"Data","Processed", out_name)
    df_subset.to_csv(out_path)


def main():
    Data_inicial,Data_agentes = get_data(filename)
    df_subset = generate_report(Data_inicial,Data_agentes)
    save_date(df_subset,filename)

if __name__ == "__main__":
    main()