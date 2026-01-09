import pandas as pd

# 1. Cargar datos desde Excel original
input_file = "DATA.xlsx"
df = pd.read_excel(input_file)

# 2. Crear dimensiones asegurando unicidad por ID

# COLABORADORES 
df_colab = df[['MATRICULA_CALIFICADO', 'ROL_CALIFICADO']].dropna()
df_colab.columns = ['ID_COLABORADOR', 'ROL_COLABORADOR']
dim_colaboradores = df_colab.drop_duplicates(subset='ID_COLABORADOR', keep='first').reset_index(drop=True)

# EVALUADORES 
df_eval = df[['MATRICULA_CALIFICADOR', 'ROL_CALIFICADOR']].dropna()
df_eval.columns = ['ID_EVALUADOR', 'ROL_EVALUADOR']
dim_evaluadores = df_eval.drop_duplicates(subset='ID_EVALUADOR', keep='first').reset_index(drop=True)

# CHAPTERS 
dim_chapters = df[['FK_CHAPTER', 'CHAPTER']].drop_duplicates().reset_index(drop=True)
dim_chapters.columns = ['ID_CHAPTER', 'NOMBRE_CHAPTER']

# ITEMS 
dim_items = df[['FK_ITEM', 'ITEM', 'TIPO_ITEM']].drop_duplicates().reset_index(drop=True)
dim_items.columns = ['ID_ITEM', 'NOMBRE_ITEM', 'TIPO_ITEM']

# 3. Crear tabla de hechos
hecho_evaluaciones = df[['FK_FECHA', 'MATRICULA_CALIFICADO', 'MATRICULA_CALIFICADOR',
                         'FK_CHAPTER', 'FK_ITEM', 'N_NIVEL', 'NIVEL']].drop_duplicates().reset_index(drop=True)

hecho_evaluaciones.columns = ['FECHA_EVALUACION', 'ID_COLABORADOR', 'ID_EVALUADOR',
                              'ID_CHAPTER', 'ID_ITEM', 'NIVEL_NUM', 'NIVEL_CAT']

# Convertir fecha
hecho_evaluaciones['FECHA_EVALUACION'] = pd.to_datetime(
    hecho_evaluaciones['FECHA_EVALUACION'].astype(str), format='%Y%m%d')


# 5. Exportar a archivo Excel con múltiples hojas
output_file = "output_ETL_final.xlsx"
with pd.ExcelWriter(output_file, engine='xlsxwriter') as writer:
    dim_colaboradores.to_excel(writer, sheet_name='dim_colaboradores', index=False)
    dim_evaluadores.to_excel(writer, sheet_name='dim_evaluadores', index=False)
    dim_chapters.to_excel(writer, sheet_name='dim_chapters', index=False)
    dim_items.to_excel(writer, sheet_name='dim_items', index=False)
    hecho_evaluaciones.to_excel(writer, sheet_name='hecho_evaluaciones', index=False)
    
print("Proceso ETL completado. Archivo guardado como output_ETL.xlsx")
