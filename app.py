import streamlit as st
import pandas as pd
from openpyxl import load_workbook
from io import BytesIO

st.title("SIS - CONTEO")

tam_bloque=4
pacientes = []

uploaded_files = st.file_uploader(
    "Sube tus archivos Excel",
    accept_multiple_files=True,
    type=['xls', 'xlsx']
)

if uploaded_files:
    for file in uploaded_files:
        df = pd.read_excel(file,skiprows=14, header=None)
        df_prueba = df[[7,8,20]]

        for i in range(0,len(df_prueba), tam_bloque):
            bloque = df_prueba.iloc[i:i+tam_bloque]
            fila = bloque.iloc[0]
            
            # Extraer datos
            edad_texto = fila[7]
            if pd.notna(edad_texto):
                try:
                    edad = int(str(edad_texto).split()[0])
                except:
                    edad = None         
            sexo = fila[8]                  
            tipo_float = fila[20]
            if pd.notna(tipo_float):
                try:
                    tipo = int(tipo_float)
                except:
                    edad = None 
            if pd.notna(sexo):
                pacientes.append({
                    "edad" : edad,
                    "sexo" : sexo,
                    "tipo" : tipo
                })

nuevo_df = pd.DataFrame(pacientes)

mujeres_p = 0
mujeres_s = 0
hombres_p = 0
hombres_s = 0

mp_0 = 0
mp_1 = 0
mp_2_4 = 0
mp_5_9 = 0
mp_10_14 = 0
mp_15_19 = 0
mp_20_29 = 0
mp_30_49 = 0
mp_50_59 = 0
mp_60 = 0

ms_0 = 0
ms_1 = 0
ms_2_4 = 0
ms_5_9 = 0
ms_10_14 = 0
ms_15_19 = 0
ms_20_29 = 0
ms_30_49 = 0
ms_50_59 = 0
ms_60 = 0

hp_0 = 0
hp_1 = 0
hp_2_4 = 0
hp_5_9 = 0
hp_10_14 = 0
hp_15_19 = 0
hp_20_29 = 0
hp_30_49 = 0
hp_50_59 = 0
hp_60 = 0

hs_0 = 0
hs_1 = 0
hs_2_4 = 0
hs_5_9 = 0
hs_10_14 = 0
hs_15_19 = 0
hs_20_29 = 0
hs_30_49 = 0
hs_50_59 = 0
hs_60 = 0


for _, row in nuevo_df.iterrows():
    sexo = row['sexo']
    tipo = row['tipo']
    edad = row['edad']

    if sexo == "MUJER":
        if tipo == 0:
            mujeres_p += 1
            
            if edad < 1:
                mp_0 += 1
            elif edad == 1:
                mp_1 += 1
            elif edad >= 2 and edad <= 4:
                mp_2_4 += 1
            elif edad >= 5 and edad <= 9:   
                mp_5_9 += 1
            elif edad >= 10 and edad <= 14:
                mp_10_14 += 1
            elif edad >= 15 and edad <= 19:
                mp_15_19 += 1
            elif edad >= 20 and edad <= 29:
                mp_20_29 += 1
            elif edad >= 30 and edad <= 49:
                mp_30_49 += 1
            elif edad >= 50 and edad <= 59:
                mp_50_59 += 1
            else:
                mp_60 += 1
                
        else:
            mujeres_s += 1
            if edad < 1:
                ms_0 += 1
            elif edad == 1:
                ms_1 += 1
            elif edad >= 2 and edad <= 4:
                ms_2_4 += 1
            elif edad >= 5 and edad <= 9:   
                ms_5_9 += 1
            elif edad >= 10 and edad <= 14:
                ms_10_14 += 1
            elif edad >= 15 and edad <= 19:
                ms_15_19 += 1
            elif edad >= 20 and edad <= 29:
                ms_20_29 += 1
            elif edad >= 30 and edad <= 49:
                ms_30_49 += 1
            elif edad >= 50 and edad <= 59:
                ms_50_59 += 1
            else:
                ms_60 += 1
                
    elif sexo == "HOMBRE":
        if tipo == 0:
            hombres_p += 1
            
            if edad < 1:
                hp_0 += 1
            elif edad == 1:
                hp_1 += 1
            elif edad >= 2 and edad <= 4:
                hp_2_4 += 1
            elif edad >= 5 and edad <= 9:   
                hp_5_9 += 1
            elif edad >= 10 and edad <= 14:
                hp_10_14 += 1
            elif edad >= 15 and edad <= 19:
                hp_15_19 += 1
            elif edad >= 20 and edad <= 29:
                hp_20_29 += 1
            elif edad >= 30 and edad <= 49:
                hp_30_49 += 1
            elif edad >= 50 and edad <= 59:
                hp_50_59 += 1
            else:
                hp_60 += 1
                
        else:
            hombres_s += 1
            if edad < 1:
                hs_0 += 1
            elif edad == 1:
                hs_1 += 1
            elif edad >= 2 and edad <= 4:
                hs_2_4 += 1
            elif edad >= 5 and edad <= 9:   
                hs_5_9 += 1
            elif edad >= 10 and edad <= 14:
                hs_10_14 += 1
            elif edad >= 15 and edad <= 19:
                hs_15_19 += 1
            elif edad >= 20 and edad <= 29:
                hs_20_29 += 1
            elif edad >= 30 and edad <= 49:
                hs_30_49 += 1
            elif edad >= 50 and edad <= 59:
                hs_50_59 += 1
            else:
                hs_60 += 1

total_pacientes = mujeres_p + mujeres_s + hombres_p + hombres_s
st.subheader("Resumen de Pacientes")
st.write(f"Total de pacientes: {total_pacientes}")
st.write(f"Mujeres primera vez: {mujeres_p}")
st.write(f"Mujeres subsecuentes: {mujeres_s}")
st.write(f"Hombres primera vez: {hombres_p}")
st.write(f"Hombres subsecuentes: {hombres_s}")


# Combinar todas las variables en un diccionario
datos = {
    "MP_0": mp_0, "MP_1": mp_1, "MP_2_4": mp_2_4, "MP_5_9": mp_5_9, "MP_10_14": mp_10_14,
    "MP_15_19": mp_15_19, "MP_20_29": mp_20_29, "MP_30_49": mp_30_49, "MP_50_59": mp_50_59, "MP_60": mp_60,
    
    "MS_0": ms_0, "MS_1": ms_1, "MS_2_4": ms_2_4, "MS_5_9": ms_5_9, "MS_10_14": ms_10_14,
    "MS_15_19": ms_15_19, "MS_20_29": ms_20_29, "MS_30_49": ms_30_49, "MS_50_59": ms_50_59, "MS_60": ms_60,
    
    "HP_0": hp_0, "HP_1": hp_1, "HP_2_4": hp_2_4, "HP_5_9": hp_5_9, "HP_10_14": hp_10_14,
    "HP_15_19": hp_15_19, "HP_20_29": hp_20_29, "HP_30_49": hp_30_49, "HP_50_59": hp_50_59, "HP_60": hp_60,
    
    "HS_0": hs_0, "HS_1": hs_1, "HS_2_4": hs_2_4, "HS_5_9": hs_5_9, "HS_10_14": hs_10_14,
    "HS_15_19": hs_15_19, "HS_20_29": hs_20_29, "HS_30_49": hs_30_49, "HS_50_59": hs_50_59, "HS_60": hs_60
}

# Convertir a DataFrame en formato largo
df_long = pd.DataFrame(list(datos.items()), columns=["Etiqueta", "Valor"])


uploaded_excel = st.file_uploader("Sube el Excel base", type=['xlsx'])

if uploaded_excel is not None:
    # Cargar libro en memoria
    wb = load_workbook(uploaded_excel)
    ws = wb.active

    # Tomar los valores del DataFrame
    valores = df_long["Valor"].tolist()

    # Escribir en la columna D (puedes cambiar la letra si quieres)
    for i, val in enumerate(valores, start=1):
        ws[f"D{i}"] = val

    # Guardar en un objeto BytesIO para descargar
    output = BytesIO()
    wb.save(output)
    output.seek(0)

    # Botón de descarga en Streamlit
    st.download_button(
        label="Descargar Excel actualizado",
        data=output,
        file_name="conteo_mes_actualizado.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    )		