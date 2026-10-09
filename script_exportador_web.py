import pandas as pd
import json
import numpy as np

def exportar_base_web():
    ruta_excel = "Base_EPS_Nacional_Estructurada_Final.xlsx"
    ruta_salida = "datos_sunass.js"
    
    print(f"Leyendo la base de datos consolidada: {ruta_excel}...")
    
    try:
        df = pd.read_excel(ruta_excel)
        df = df.replace({np.nan: None})
        
        # 1. Dataset de Metas e Indicadores (Tablero y Reportes)
        columnas_metas = [
            'Cod_EP', 'Nombre_EPS', 'Periodo_Regulatorio', 'Departamento', 'Tamano_EPS', 'Estado_Evaluacion',
            'ID_Meta_General', 'Meta_Principal', 'Nivel_Evaluacion', 'Descripcion_Especifica', 'Unidad_Medida',
            'Etapa', 'Año_Regulatorio', 'Valor_Exigido', 'Valor_Ejecutado', 'ICI'
        ]
        cols_metas_presentes = [c for c in columnas_metas if c in df.columns]
        registros_metas = df[cols_metas_presentes].to_dict(orient='records')
        
        # 2. Matriz Consolidada de Informes por Año y Etapa
        cols_informes = [
            'Cod_EP', 'Nombre_EPS', 'Periodo_Regulatorio', 'Año_Regulatorio', 'Estado_Evaluacion',
            'Inf_Fisc_Inicial', 'Inf_Fisc_Inicial_Link',
            'Inf_Fisc_Final', 'Inf_Fisc_Final_Link',
            'Inf_Fiscalizacion', 'Inf_Fiscalizacion_Link',
            'Inf_Fisc_Complementario', 'Inf_Fisc_Complementario_Link',
            'Inf_Fisc_MedidasCorrectivas', 'Inf_Fisc_MedidasCorrectivas_Link',
            'Inf_PAS_Instruccion', 'Inf_PAS_Instruccion_Link',
            'Inf_PAS_Complementario', 'Inf_PAS_Complementario_Link',
            'Inf_PAS_Decision', 'Inf_PAS_Decision_Link',
            'Recurso_Reconsideracion', 'Recurso_Reconsideracion_Link',
            'Recurso_Apelacion', 'Recurso_Apelacion_Link'
        ]
        cols_inf_presentes = [c for c in cols_informes if c in df.columns]
        df_inf = df[cols_inf_presentes].drop_duplicates(subset=['Cod_EP', 'Periodo_Regulatorio', 'Año_Regulatorio'])
        registros_informes = df_inf.to_dict(orient='records')
        
        print(f"Escribiendo {len(registros_metas)} metas y {len(registros_informes)} registros de informes en '{ruta_salida}'...")
        with open(ruta_salida, "w", encoding="utf-8") as f:
            f.write("// Base de datos generada automáticamente desde Python\n")
            f.write("const baseDatosNacional = ")
            json.dump(registros_metas, f, ensure_ascii=False, separators=(',', ':'))
            f.write(";\n\n")
            f.write("// Matriz de informes y resoluciones linkeados por EPS, Periodo y Año\n")
            f.write("const matrizInformesNacional = ")
            json.dump(registros_informes, f, ensure_ascii=False, separators=(',', ':'))
            f.write(";\n")
            
        print(f"¡Éxito! Se generó '{ruta_salida}' con soporte completo para la Matriz de Informes.")
        
    except FileNotFoundError:
        print(f"Error: No se encontró el archivo '{ruta_excel}'.")
    except Exception as e:
        print(f"Ocurrió un error inesperado: {e}")

if __name__ == "__main__":
    exportar_base_web()