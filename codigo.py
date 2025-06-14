import pandas as pd
import os
import numpy as np
import tkinter as tk
from tkinter import filedialog, messagebox
import glob

def seleccionar_carpeta_entrada():
    """Permite al usuario seleccionar la carpeta con los archivos Excel"""
    root = tk.Tk()
    root.withdraw()  # Ocultar la ventana principal
    
    carpeta = filedialog.askdirectory(
        title="Selecciona la carpeta con los archivos Excel a procesar"
    )
    
    if not carpeta:
        messagebox.showinfo("Cancelado", "No se seleccionó ninguna carpeta.")
        return None
    
    # Buscar todos los archivos .xlsx en la carpeta seleccionada
    patron_excel = os.path.join(carpeta, "*.xlsx")
    archivos_excel = glob.glob(patron_excel)
    
    if not archivos_excel:
        messagebox.showwarning("Sin archivos", f"No se encontraron archivos Excel (.xlsx) en:\n{carpeta}")
        return None
    
    print(f"\n📁 Carpeta seleccionada: {carpeta}")
    print(f"📋 Archivos encontrados ({len(archivos_excel)}):")
    for archivo in archivos_excel:
        print(f"  • {os.path.basename(archivo)}")
    
    return archivos_excel

def seleccionar_archivo_salida():
    """Permite al usuario seleccionar dónde guardar el archivo de salida"""
    root = tk.Tk()
    root.withdraw()  # Ocultar la ventana principal
    
    archivo_salida = filedialog.asksaveasfilename(
        title="Guardar archivo procesado como...",
        defaultextension=".xlsx",
        filetypes=[("Archivos Excel", "*.xlsx"), ("Todos los archivos", "*.*")],
        initialfile="archivo_procesado.xlsx"
    )
    
    if not archivo_salida:
        messagebox.showinfo("Cancelado", "No se seleccionó ubicación de salida.")
        return None
    
    return archivo_salida

# Seleccionar archivos de entrada
print("🔍 Selecciona la carpeta con los archivos Excel...")
rutas_excel = seleccionar_carpeta_entrada()

if not rutas_excel:
    print("❌ Proceso cancelado. No se seleccionaron archivos.")
    exit()

# Seleccionar archivo de salida
print("\n💾 Selecciona dónde guardar el archivo procesado...")
ruta_salida = seleccionar_archivo_salida()

if not ruta_salida:
    print("❌ Proceso cancelado. No se seleccionó archivo de salida.")
    exit()

print(f"\n✅ Configuración completada:")
print(f"📥 Archivos a procesar: {len(rutas_excel)}")
print(f"📤 Archivo de salida: {os.path.basename(ruta_salida)}")
print(f"📍 Ubicación: {os.path.dirname(ruta_salida)}")

def procesar_archivo(ruta_excel):
    print(f"\nProcesando archivo: {os.path.basename(ruta_excel)}")
    try:
        df = pd.read_excel(ruta_excel)

        # Normalizar nombres de columnas
        df.columns = [str(col).strip().lower() for col in df.columns]

        # Verificar y crear columnas necesarias si no existen
        columnas_requeridas = [
            "n.º de elemento", "material", "nombre", "configuracion", "cantidad",
            "espesor", "canto", "ancho", "profundidad", "tipo"
        ]
        for col in columnas_requeridas:
            if col not in df.columns:
                df[col] = np.nan

        # Convertir columnas a tipos apropiados
        df["material"] = df["material"].astype(str)
        df["nombre"] = df["nombre"].astype(str)
        df["configuracion"] = df["configuracion"].astype(str)
        df["canto"] = df["canto"].apply(lambda x: "" if pd.isna(x) else str(x))

        # Limpieza de texto
        df["material_limpio"] = df["material"].str.strip().str.lower()
        df["tipo_limpio"] = df["tipo"].astype(str).str.strip().str.lower()

        df_resina = pd.DataFrame()
        df_okuome = pd.DataFrame()

        for i, row in df.iterrows():
            material_actual = row["material_limpio"]
            
            # Procesar tanto resina de melamina como okuome
            if material_actual in ["resina de melamina", "okuome"]:
                nueva_fila = {
                    "material": row["material"],
                    "nombre": row["nombre"],
                    "configuracion": row["configuracion"],
                    "ancho": row["ancho"],
                    "profundidad": row["profundidad"],
                    "espesor": row["espesor"],
                    "canto": row["canto"] if pd.notna(row["canto"]) else ""
                }

                cantidad_actual = float(row["cantidad"]) if pd.notna(row["cantidad"]) else 1

                # Buscar ensamblaje más cercano hacia arriba
                ensamblaje_idx = None
                for j in range(i - 1, -1, -1):
                    if df.loc[j, "tipo_limpio"] == "ensamblaje":
                        ensamblaje_idx = j
                        break

                inicio_busqueda = ensamblaje_idx + 1 if ensamblaje_idx is not None else 0

                cantidad_subensamblaje = 1
                for j in range(i - 1, inicio_busqueda - 1, -1):
                    if df.loc[j, "tipo_limpio"] == "subensamblaje":
                        if pd.notna(df.loc[j, "cantidad"]):
                            cantidad_subensamblaje = float(df.loc[j, "cantidad"])
                        break

                cantidad_ensamblaje = 1
                if ensamblaje_idx is not None and pd.notna(df.loc[ensamblaje_idx, "cantidad"]):
                    cantidad_ensamblaje = float(df.loc[ensamblaje_idx, "cantidad"])

                nueva_fila["cantidad_calculada"] = cantidad_actual * cantidad_ensamblaje * cantidad_subensamblaje

                # Asignar etiquetas según el material
                if material_actual == "resina de melamina":
                    if pd.notna(row["espesor"]):
                        if row["espesor"] == 18:
                            nueva_fila["etiqueta"] = "melamina gris"
                        elif row["espesor"] == 5:
                            nueva_fila["etiqueta"] = "negro"
                        else:
                            nueva_fila["etiqueta"] = ""
                    else:
                        nueva_fila["etiqueta"] = ""
                elif material_actual == "okuome":
                    # Agregar lógica específica para okuome si es necesaria
                    nueva_fila["etiqueta"] = "okuome"

                nueva_fila["nombre_especifico"] = f"{nueva_fila['configuracion']}_{nueva_fila['nombre']}_{nueva_fila['etiqueta']}"
                nueva_fila["archivo_origen"] = os.path.basename(ruta_excel)

                # Agregar a DataFrame correspondiente
                if material_actual == "resina de melamina":
                    df_resina = pd.concat([df_resina, pd.DataFrame([nueva_fila])], ignore_index=True)
                elif material_actual == "okuome":
                    df_okuome = pd.concat([df_okuome, pd.DataFrame([nueva_fila])], ignore_index=True)

        return df_resina, df_okuome

    except Exception as e:
        print(f"Error al procesar {os.path.basename(ruta_excel)}: {str(e)}")
        return pd.DataFrame(), pd.DataFrame()

def crear_resumenes(df_material, nombre_material):
    """Crear los diferentes resúmenes para un material específico"""
    if df_material.empty:
        return None, None, None, None
    
    # Resumen por nombre específico
    nombres_unicos = df_material["nombre_especifico"].unique()
    datos_resumen = []
    for nombre in nombres_unicos:
        datos_filtrados = df_material[df_material["nombre_especifico"] == nombre]
        cantidad_total = datos_filtrados["cantidad_calculada"].sum()
        primera_fila = datos_filtrados.iloc[0]
        datos_resumen.append({
            "nombre": nombre,
            "cantidad_calculada": cantidad_total,
            "ancho": primera_fila["ancho"],
            "profundidad": primera_fila["profundidad"],
            "espesor": primera_fila["espesor"],
            "canto": primera_fila["canto"]
        })
    resumen_etiqueta = pd.DataFrame(datos_resumen)

    # Resumen por archivo
    archivos_unicos = df_material["archivo_origen"].unique()
    datos_archivo = []
    for archivo in archivos_unicos:
        datos_filtrados = df_material[df_material["archivo_origen"] == archivo]
        cantidad_total = datos_filtrados["cantidad_calculada"].sum()
        datos_archivo.append({
            "archivo_origen": archivo,
            "cantidad_calculada": cantidad_total
        })
    resumen_archivo = pd.DataFrame(datos_archivo)

    # Detalle completo
    resultado = df_material[["cantidad_calculada", "ancho", "profundidad", "espesor", "canto", "nombre_especifico", "archivo_origen"]]
    resultado = resultado.rename(columns={"nombre_especifico": "nombre"})

    # Query agrupado
    df_query = resultado[resultado["nombre"].notna()]
    df_query_agrupado = df_query.groupby(["ancho", "profundidad", "espesor", "canto"], dropna=False)\
                                .agg({"cantidad_calculada": "sum"}).reset_index()
    df_query_agrupado = df_query_agrupado.rename(columns={"cantidad_calculada": "Total"})
    df_query_agrupado = df_query_agrupado[["Total", "ancho", "profundidad", "espesor", "canto"]]

    return resultado, resumen_etiqueta, resumen_archivo, df_query_agrupado

# Procesar todos los archivos
resultados_resina = []
resultados_okuome = []

for ruta in rutas_excel:
    try:
        df_resina, df_okuome = procesar_archivo(ruta)
        
        if not df_resina.empty:
            resultados_resina.append(df_resina)
            print(f"✓ Resina procesada: {os.path.basename(ruta)}")
        
        if not df_okuome.empty:
            resultados_okuome.append(df_okuome)
            print(f"✓ Okuome procesado: {os.path.basename(ruta)}")
            
        if df_resina.empty and df_okuome.empty:
            print(f"⚠️ Sin resultados para: {os.path.basename(ruta)}")
            
    except Exception as e:
        print(f"❌ Error: {os.path.basename(ruta)} -> {str(e)}")

# Combinar y procesar resultados
try:
    with pd.ExcelWriter(ruta_salida, engine="openpyxl") as writer:
        hojas_creadas = []
        
        # Procesar resina de melamina
        if resultados_resina:
            df_resina_combinado = pd.concat(resultados_resina, ignore_index=True)
            resultado_resina, resumen_etiqueta_resina, resumen_archivo_resina, query_resina = crear_resumenes(df_resina_combinado, "resina")
            
            resultado_resina.to_excel(writer, sheet_name="Detalle Resina", index=False)
            resumen_etiqueta_resina.to_excel(writer, sheet_name="Resumen Nombre Resina", index=False)
            resumen_archivo_resina.to_excel(writer, sheet_name="Resumen Archivo Resina", index=False)
            query_resina.to_excel(writer, sheet_name="Query Resina", index=False)
            
            hojas_creadas.extend(["Detalle Resina", "Resumen Nombre Resina", "Resumen Archivo Resina", "Query Resina"])
            print("✓ Hojas de resina de melamina creadas")

        # Procesar okuome
        if resultados_okuome:
            df_okuome_combinado = pd.concat(resultados_okuome, ignore_index=True)
            resultado_okuome, resumen_etiqueta_okuome, resumen_archivo_okuome, query_okuome = crear_resumenes(df_okuome_combinado, "okuome")
            
            resultado_okuome.to_excel(writer, sheet_name="Detalle Okuome", index=False)
            resumen_etiqueta_okuome.to_excel(writer, sheet_name="Resumen Nombre Okuome", index=False)
            resumen_archivo_okuome.to_excel(writer, sheet_name="Resumen Archivo Okuome", index=False)
            query_okuome.to_excel(writer, sheet_name="Query Okuome", index=False)
            
            hojas_creadas.extend(["Detalle Okuome", "Resumen Nombre Okuome", "Resumen Archivo Okuome", "Query Okuome"])
            print("✓ Hojas de okuome creadas")

        # Crear hoja combinada si hay ambos materiales
        if resultados_resina and resultados_okuome:
            df_combinado_total = pd.concat([df_resina_combinado, df_okuome_combinado], ignore_index=True)
            resultado_total = df_combinado_total[["cantidad_calculada", "ancho", "profundidad", "espesor", "canto", "nombre_especifico", "archivo_origen", "material"]]
            resultado_total = resultado_total.rename(columns={"nombre_especifico": "nombre"})
            resultado_total.to_excel(writer, sheet_name="Detalle Combinado", index=False)
            hojas_creadas.append("Detalle Combinado")
            print("✓ Hoja combinada creada")

    print(f"\n✅ Exportación exitosa. Hojas creadas: {len(hojas_creadas)}")
    print(f"📋 Hojas: {', '.join(hojas_creadas)}")
    print(f"📁 Archivo guardado en:\n{ruta_salida}")

except Exception as e:
    print(f"\n❌ Error al crear el archivo de salida: {str(e)}")

if not resultados_resina and not resultados_okuome:
    print("\n⚠️ No se encontraron resultados para exportar.")