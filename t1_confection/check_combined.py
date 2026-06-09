import pandas as pd
df = pd.read_csv('OSTRAM_Combined_Inputs_Outputs.csv', usecols=['Scenario','TotalCapacityAnnual','ProductionByTechnologyAnnual','TotalDiscountedCost'], low_memory=False)
for s in sorted(df['Scenario'].unique()):
    sub = df[df['Scenario']==s]
    cap = sub['TotalCapacityAnnual'].sum()
    gen = sub['ProductionByTechnologyAnnual'].sum()
    cost = sub['TotalDiscountedCost'].sum()
    print(f"{s}: cap={cap:,.0f}  gen={gen:,.0f}  cost={cost:,.0f}")
