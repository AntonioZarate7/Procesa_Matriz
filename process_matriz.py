import pandas as pd

# Cargar archivo
#file_path = "Copia de Matiz Emision - RP 2025 No afectacion.xlsx"
file_path = "Matriz de emisión RP 2026_correccion en CEDA y GURA.xlsx"
df = pd.read_excel(file_path, sheet_name="Simplificada")

# Limpieza inicial
df_clean = df[df["ID"].notna()].copy()
df_clean["ID"] = df_clean["ID"].astype(int)

# Número de asegurados por ID
df_num_asegurados = df_clean["ID"].value_counts().reset_index()
df_num_asegurados.columns = ["TCID", "NUM ASEGURADOS"]
df_num_asegurados = df_num_asegurados.sort_values("TCID").reset_index(drop=True)

# Prima esperada
df_clean["PRIMA ESPERADA"] = df_clean["Pma + Der"]
df_prima_esperada = df_clean.groupby("ID", as_index=False)["PRIMA ESPERADA"].sum()

# Riesgo
def riesgo_por_id(grupo):
    riesgos = grupo.sort_index()["Riesgo"].fillna("").str.upper().values
    titular = riesgos[0] == "PREFERENTE"
    asegurados = "PREFERENTE" in riesgos[1:]
    if titular and asegurados: return "Ambos"
    elif titular: return "Titular"
    elif asegurados: return "Asegurado"
    return "NO"

df_riesgo = df_clean.groupby("ID").apply(riesgo_por_id).reset_index()
df_riesgo.columns = ["TCID", "RIESGO"]

# Tipo de Deducible
col_deducible = [c for c in df_clean.columns if "Tipo de Deducible" in c][0]
df_tipo_deducible = df_clean.groupby("ID")[col_deducible].first().reset_index()
df_tipo_deducible.columns = ["TCID", "TIPO DE DEDUCIBLE"]

# Tipo de Coaseguro
col_coaseguro = [c for c in df_clean.columns if "Tipo de Coaseguro" in c][0]
df_tipo_coaseguro = df_clean.groupby("ID")[col_coaseguro].first().reset_index()
df_tipo_coaseguro.columns = ["TCID", "TIPO DE COASEGURO"]

# Plan y Moneda
df_plan = df_clean.groupby("ID")["Plan"].first().reset_index()
df_plan.columns = ["TCID", "PLAN"]
df_plan["PLAN"] = df_plan["PLAN"].replace({
    "AMF A": "ALFA MEDICAL FLEX A",
    "AMF B": "ALFA MEDICAL FLEX B",
    "AMI": "ALFA MEDICAL INTERNACIONAL"
})
df_plan["MONEDA"] = df_plan["PLAN"].apply(lambda x: "DOLAR" if x == "ALFA MEDICAL INTERNACIONAL" else "PESO")

# FORMA DE PAGO
df_forma_pago = df_clean.groupby("ID")["Forma de pago"].first().reset_index()
df_forma_pago.columns = ["TCID", "FORMA DE PAGO"]

# Deducible + Exceso
df_deducible = df_clean.groupby("ID")["Deducible"].first().reset_index()
df_deducible.columns = ["TCID", "DEDUCIBLE"]
df_deducible["DEDUCIBLE EN EXCESO"] = df_deducible["DEDUCIBLE"].apply(
    lambda x: "SI" if pd.notnull(x) and x >= 200000 else "NO"
)
df_deducible["DEDUCIBLE"] = df_deducible["DEDUCIBLE"].apply(lambda x: str(int(x)) if pd.notnull(x) else "")

# Coaseguro
df_coaseguro = df_clean.groupby("ID")["Coaseguro"].first().reset_index()
df_coaseguro.columns = ["TCID", "COASEGURO"]
df_coaseguro["COASEGURO"] = df_coaseguro["COASEGURO"].apply(lambda x: int(x * 100) if pd.notnull(x) else "")

# GURA
df_gura = df_clean.groupby("ID")["Incremento GURA"].first().reset_index()
df_gura.columns = ["TCID", "GURA"]
df_gura["GURA"] = df_gura["GURA"].apply(lambda x: int(float(str(x).replace('%', '')) * 100) if pd.notnull(x) else "")

# Suma Asegurada
df_suma = df_clean.groupby("ID")["Suma Asegurada"].first().reset_index()
df_suma.columns = ["TCID", "SUMA ASEGURADA"]
df_suma["SUMA ASEGURADA"] = df_suma["SUMA ASEGURADA"].apply(lambda x: str(int(x)) if pd.notnull(x) else "")

# CPF → TF/GF
df_cpf = df_clean.groupby("ID")["CPF"].first().reset_index()
df_cpf.columns = ["TCID", "CPF"]
df_cpf["TF"] = df_cpf["CPF"].apply(lambda x: "TF" if pd.notnull(x) and str(x).strip() else "NO")
df_cpf["GF"] = df_cpf["CPF"].apply(lambda x: "GF" if pd.notnull(x) and str(x).strip() else "NO")
df_cpf.drop(columns=["CPF"], inplace=True)

# CAE
df_cae = df_clean.groupby("ID")["CAE"].first().reset_index()
df_cae.columns = ["TCID", "CAE"]
df_cae["CAE"] = df_cae["CAE"].apply(lambda x: "CAE" if pd.notnull(x) and str(x).strip() else "NO")

#CEC
df_cec = df_clean.groupby("ID")["CEC"].first().reset_index()
df_cec.columns = ["TCID", "CEC"]
df_cec["CEC"] = df_cec["CEC"].apply(lambda x: "CEC" if pd.notnull(x) and str(x).strip() else "NO")

#CEE
df_cee = df_clean.groupby("ID")["CEE"].first().reset_index()
df_cee.columns = ["TCID", "CEE"]
df_cee["CEE"] = df_cee["CEE"].apply(lambda x: "CEE" if pd.notnull(x) and str(x).strip() else "NO")

#CEDA
df_ceda = df_clean.groupby("ID")["CEDA"].first().reset_index()
df_ceda.columns = ["TCID", "CEDA"]
df_ceda["CEDA"] = df_ceda["CEDA"].apply(lambda x: "CEDA" if pd.notnull(x) and str(x).strip() else "NO")

#DENTAL
df_dental = df_clean.groupby("ID")["DENTAL"].first().reset_index()
df_dental.columns = ["TCID", "DENTAL"]
df_dental["DENTAL"] = df_dental["DENTAL"].apply(lambda x: "DP" if pd.notnull(x) and str(x).strip() else "NO")

#CEDA PREMIUM
df_cedap = df_clean.groupby("ID")["CEDA PREM"].first().reset_index()
df_cedap.columns = ["TCID", "CEDA PREM"]
df_cedap["CEDA PREM"] = df_cedap["CEDA PREM"].apply(lambda x: "CEDAP" if pd.notnull(x) and str(x).strip() else "NO")

#CRFCA
df_crfca = df_clean.groupby("ID")["CRFCA"].first().reset_index()
df_crfca.columns = ["TCID", "CRFCA"]
df_crfca["CRFCA"] = df_crfca["CRFCA"].apply(lambda x: "CRFCA" if pd.notnull(x) and str(x).strip() else "NO")

# AMCD: AMCDC + AMCDA1-10
df_contratantes = df_clean.groupby("ID", as_index=False).first()
df_amcdc = df_contratantes[["ID", "AMCD"]].rename(columns={"ID": "TCID"})
df_amcdc["AMCDC"] = df_amcdc["AMCD"].apply(lambda x: "AMCD" if pd.notnull(x) and str(x).strip() else "NO")
df_amcdc.drop(columns=["AMCD"], inplace=True)

df_amcd_multi = pd.DataFrame()
for i in range(1, 11):
    dfn = df_clean.groupby("ID").nth(i).reset_index()[["ID", "AMCD"]].rename(columns={"ID": "TCID"})
    dfn[f"AMCDA{i}"] = dfn["AMCD"].apply(lambda x: "AMCD" if pd.notnull(x) and str(x).strip() else "NO")
    dfn.drop(columns=["AMCD"], inplace=True)
    df_amcd_multi = dfn if df_amcd_multi.empty else df_amcd_multi.merge(dfn, on="TCID", how="outer")
df_amcd_multi.fillna("NO", inplace=True)
df_amcd_all = pd.merge(df_amcdc, df_amcd_multi, on="TCID", how="outer")

# CETTE: CETTEC + CETTEA1-A10
df_cettec = df_contratantes[["ID", "CETTE"]].rename(columns={"ID": "TCID"})
df_cettec["CETTEC"] = df_cettec["CETTE"].apply(lambda x: "CETTE" if pd.notnull(x) and str(x).strip() else "NO")
df_cettec.drop(columns=["CETTE"], inplace=True)

df_cette_multi = pd.DataFrame()
for i in range(1, 11):
    dfn = df_clean.groupby("ID").nth(i).reset_index()[["ID", "CETTE"]].rename(columns={"ID": "TCID"})
    dfn[f"CETTEA{i}"] = dfn["CETTE"].apply(lambda x: "CETTE" if pd.notnull(x) and str(x).strip() else "NO")
    dfn.drop(columns=["CETTE"], inplace=True)
    df_cette_multi = dfn if df_cette_multi.empty else df_cette_multi.merge(dfn, on="TCID", how="outer")
df_cette_multi.fillna("NO", inplace=True)
df_cette_all = pd.merge(df_cettec, df_cette_multi, on="TCID", how="outer")


# DNIC (Nombre del contratante)
df_contratantes = df_clean.groupby("ID", as_index=False).first()
df_dnic = df_contratantes[["ID", "Nombre "]].rename(columns={"ID": "TCID", "Nombre ": "DNIC"})

# DNIA1 a DNIA10 (asegurados)
df_dnia_multi = pd.DataFrame()

for i in range(1, 11):
    df_ni = (
        df_clean.groupby("ID")
        .nth(i)
        .reset_index()[["ID", "Nombre "]].rename(columns={"ID": "TCID"})
    )
    col_name = f"DNIA{i}"
    df_ni[col_name] = df_ni["Nombre "]
    df_ni.drop(columns=["Nombre "], inplace=True)

    df_dnia_multi = df_ni if df_dnia_multi.empty else df_dnia_multi.merge(df_ni, on="TCID", how="outer")

# Llenar vacíos con cadena vacía
df_dnia_multi.fillna("", inplace=True)

# Unir DNIC + DNIAx
df_dni_final = pd.merge(df_dnic, df_dnia_multi, on="TCID", how="outer")


# CMAC (contratante: primer registro por ID)
df_contratantes = df_clean.groupby("ID", as_index=False).first()
df_cmac = df_contratantes[["ID", "Edad", "Sexo"]].rename(columns={"ID": "TCID"})
df_cmac["CMAC"] = df_cmac.apply(
    lambda row: "SI" if 15 <= row["Edad"] <= 44 and str(row["Sexo"]).strip().upper() == "F" else "NO",
    axis=1
)
df_cmac = df_cmac[["TCID", "CMAC"]]

# CMAA1 a CMAA10 (asegurados)
df_cmaa_multi = pd.DataFrame()

for i in range(1, 11):
    df_ni = df_clean.groupby("ID").nth(i).reset_index()[["ID", "Edad", "Sexo"]].rename(columns={"ID": "TCID"})
    col_name = f"CMAA{i}"
    df_ni[col_name] = df_ni.apply(
        lambda row: "SI" if 15 <= row["Edad"] <= 44 and str(row["Sexo"]).strip().upper() == "F" else "NO",
        axis=1
    )
    df_ni = df_ni[["TCID", col_name]]
    df_cmaa_multi = df_ni if df_cmaa_multi.empty else df_cmaa_multi.merge(df_ni, on="TCID", how="outer")

# Rellenar vacíos con "NO"
df_cmaa_multi.fillna("NO", inplace=True)

# Unir CMAC y CMAA1-10
df_cma_final = pd.merge(df_cmac, df_cmaa_multi, on="TCID", how="outer")


# Consolidar todos los datos en el orden deseado
df_resultado = df_num_asegurados
df_resultado = df_resultado.merge(df_prima_esperada, left_on="TCID", right_on="ID").drop(columns=["ID"])
df_resultado["FLUJO"] = "emision"
df_resultado["TIPO"] = "GMM"
df_resultado["RUN"] = 0
df_resultado["CLAVE"] = "CG-111-X"
df_resultado = df_resultado.merge(df_riesgo, on="TCID")
df_resultado = df_resultado.merge(df_tipo_deducible, on="TCID")
df_resultado = df_resultado.merge(df_tipo_coaseguro, on="TCID")
df_resultado = df_resultado.merge(df_plan, on="TCID")
df_resultado["FECHA SOLICITUD"] = "03032025"
df_resultado["FECHA EFECTO"] = "03032025"
df_resultado["CONDUCTO DE COBRO"] = "AGENTE"
df_resultado = df_resultado.merge(df_forma_pago, on="TCID")  # <-- Incluye FORMA DE PAGO
df_resultado["AGENTE"] = "11026"  # <-- AGENTE valor fijo
df_resultado["PARTICIPACION"] = "100"
df_resultado = df_resultado.merge(df_deducible, on="TCID")
df_resultado = df_resultado.merge(df_coaseguro, on="TCID")
df_resultado = df_resultado.merge(df_gura, on="TCID")
df_resultado = df_resultado.merge(df_suma, on="TCID")
df_resultado = df_resultado.merge(df_cpf, on="TCID")
df_resultado = df_resultado.merge(df_cae, on="TCID")
df_resultado = df_resultado.merge(df_cec, on="TCID")
df_resultado = df_resultado.merge(df_cee, on="TCID")
df_resultado = df_resultado.merge(df_ceda, on="TCID")
df_resultado = df_resultado.merge(df_dental, on="TCID")
df_resultado = df_resultado.merge(df_cedap, on="TCID")
df_resultado = df_resultado.merge(df_crfca, on="TCID")
#df_resultado = df_resultado.merge(df_amcdc, on="TCID")
#df_resultado = df_resultado.merge(df_a1, on="TCID", how="left")
df_resultado = df_resultado.merge(df_amcd_all, on="TCID", how="left")
df_resultado = df_resultado.merge(df_cette_all, on="TCID", how="left")
df_resultado = df_resultado.merge(df_dni_final, on="TCID", how="left")
df_resultado = df_resultado.merge(df_cma_final, on="TCID", how="left")
df_resultado["POLIZA"] = ""

# Reordenar: FORMA DE PAGO antes de AGENTE
"""cols = df_resultado.columns.tolist()
if "AGENTE" in cols and "FORMA DE PAGO" in cols:
    cols.insert(cols.index("AGENTE"), cols.pop(cols.index("FORMA DE PAGO")))
    df_resultado = df_resultado[cols]"""

# Exportar
df_resultado.to_csv("matriz_emision.csv", index=False)
print("✅ Archivo 'matriz_emision.csv' generado con exito.")
