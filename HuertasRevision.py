import pandas as pd
import glob
import os

# --- Configuration ---
# Paths
path_resultados = r"H:\Mi unidad\Modernización_Educativa\Gerencia de Evaluación_Proyectos_Análisis\PAARS_Warehouse\2026\01_Resultados_Febrero\Interim_CSVs\Resultados\*.csv"
path_progreso = r"H:\Mi unidad\Modernización_Educativa\Gerencia de Evaluación_Proyectos_Análisis\PAARS_Warehouse\2026\02_PROGRESO_Abril\Interim_CSVs\Resultados\*.csv"

# Target variables
target_schools = [11963, 86389]

# Column names
student_id_col = 'Documento' 
grade_col = 'Grado'  # <-- UPDATE THIS to the exact column name in your CSV

# Grades to filter (Update this list to match exactly how grades are written in your file)
# Example: If written as text, change to ['8°', '9°', '10°'] or ['8', '9', '1er año']
target_grades = [8, 9, 10] 

def count_students_both_subjects_any_score(file_pattern, test_name):
    all_files = glob.glob(file_pattern)
    mat_list = []
    lec_list = []
    
    for file in all_files:
        file_name = os.path.basename(file).upper()
        
        try:
            # Read each CSV file
            df = pd.read_csv(file, low_memory=False)
            
            # Verify the necessary columns exist before proceeding
            missing_cols = [col for col in ['Nro de centro', student_id_col, grade_col] if col not in df.columns]
            if missing_cols:
                print(f"Skipping {os.path.basename(file)}: Missing columns {missing_cols}")
                continue
            
            # Filter for the target schools AND target grades
            school_filter = df['Nro de centro'].isin(target_schools)
            grade_filter = df[grade_col].isin(target_grades)
            
            df_filtered = df[school_filter & grade_filter].copy()
            
            # Sort into Math or Lengua lists based on filename
            if 'MAT' in file_name:
                mat_list.append(df_filtered)
            elif 'LEC' in file_name:
                lec_list.append(df_filtered)
                
        except Exception as e:
            print(f"Error reading {os.path.basename(file)}: {e}")
            
    # If a folder doesn't have both types of files, we can't do an intersection
    if not mat_list or not lec_list:
        return pd.DataFrame()

    # Combine all Math files and all Lengua files into two separate master DataFrames
    df_mat = pd.concat(mat_list, ignore_index=True)
    df_lec = pd.concat(lec_list, ignore_index=True)

    # Keep only School ID and Student ID columns, and remove duplicates just in case
    df_mat_unique = df_mat[['Nro de centro', student_id_col]].drop_duplicates()
    df_lec_unique = df_lec[['Nro de centro', student_id_col]].drop_duplicates()

    # Merge the two tables using an 'inner' join. 
    # This only keeps students who exist in BOTH the Math table AND the Lengua table.
    df_both = pd.merge(df_mat_unique, df_lec_unique, on=['Nro de centro', student_id_col], how='inner')
    
    # Group by School and count the remaining rows
    counts = df_both.groupby('Nro de centro').size().reset_index(name=f'Students in Both Subjects (Grades {target_grades})')
    counts['Test'] = test_name
    
    return counts

# --- Execution ---

print("Processing Prueba de Resultados...")
resultados_counts = count_students_both_subjects_any_score(path_resultados, "Prueba de Resultados")

print("\nProcessing Prueba de Progreso 1...")
progreso_counts = count_students_both_subjects_any_score(path_progreso, "Prueba de Progreso 1")

# Combine the results from both exams
final_results = pd.concat([resultados_counts, progreso_counts], ignore_index=True)

if not final_results.empty:
    print("\n--- Final Counts ---")
    print(final_results.to_string(index=False))
else:
    print("\nNo students found matching the criteria (or files were missing).")