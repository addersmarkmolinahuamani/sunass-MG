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
        
        # Columnas optimizadas para visualización web y reportes ad-hoc
        columnas_web = [
            'Cod_EP', 'Nombre_EPS', 'Periodo_Regulatorio', 'Departamento', 'Tamano_EPS', 'Estado_Evaluacion',
            'ID_Meta_General', 'Meta_Principal', 'Nivel_Evaluacion', 'Descripcion_Especifica', 'Unidad_Medida',
            'Etapa', 'Año_Regulatorio', 'Valor_Exigido', 'Valor_Ejecutado', 'ICI',
            'Informe_Etapa', 'Informe_Etapa_Link',
            'Inf_Fiscalizacion', 'Inf_Fiscalizacion_Link',
            'Inf_PAS_Decision', 'Inf_PAS_Decision_Link'
        ]
        
        # Filtrar solo las columnas existentes
        cols_presentes = [c for c in columnas_web if c in df.columns]
        df_web = df[cols_presentes]
        
        registros = df_web.to_dict(orient='records')
        
        print(f"Escribiendo {len(registros)} registros en formato compacto en '{ruta_salida}'...")
        with open(ruta_salida, "w", encoding="utf-8") as f:
            f.write("// Base de datos generada automáticamente desde Python\n")
            f.write("const baseDatosNacional = ")
            json.dump(registros, f, ensure_ascii=False, separators=(',', ':'))
            f.write(";\n")
            
        print(f"¡Éxito! Se generó '{ruta_salida}' optimizado para GitHub Pages con enlaces de informes.")
        
    except FileNotFoundError:
        print(f"Error: No se encontró el archivo '{ruta_excel}'.")
    except Exception as e:
        print(f"Ocurrió un error inesperado: {e}")

if __name__ == "__main__":
    exportar_base_web()