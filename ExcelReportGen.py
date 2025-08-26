import pandas as pd

df = pd.read_excel("sale_data.xlsx")
summary = df.groupby("Region")["Amount"].sum()
summary.to_excel("summary_report.xlsx")