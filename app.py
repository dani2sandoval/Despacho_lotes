import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Control de Rotación IGSS", page_icon="📦", layout="centered")

st.title("📦 Control de Rotación de Inventario (FEFO)")
st.subheader("Hospital General IGSS Jalapa")
st.markdown("Sube los archivos exportados del sistema para generar el despacho automático de lotes.")

# --- SECCIÓN DE CARGA DE ARCHIVOS ---
st.sidebar.header("📁 Carga de Documentos")
archivo_bodega = st.sidebar.file_uploader("1. Inventario de Bodega (.xlsx)", type=["xlsx"])
archivo_farmacia = st.sidebar.file_uploader("2. Inventario de Farmacia (.xlsx)", type=["xlsx"])
archivo_solicitud = st.sidebar.file_uploader("3. Requisición/Solicitud de Farmacia (.xlsx)", type=["xlsx"])


def optimizar_dataframe(file, palabras_clave):
    df_sucio = pd.read_excel(file, header=None)
    fila_encabezado = 0
    for i, fila in df_sucio.iterrows():
        fila_str = [str(x).strip().upper() for x in fila.dropna()]
        if any(any(p in col for p in palabras_clave) for col in fila_str):
            fila_encabezado = i
            break

    df_limpio = pd.read_excel(file, skiprows=fila_encabezado)
    nuevos_columnas = []
    for col in df_limpio.columns:
        col_str = str(col).strip().upper()
        if "UNNAMED" in col_str:
            nuevos_columnas.append("VACIA_COMBINADA")
        else:
            nuevos_columnas.append(col_str)
    df_limpio.columns = nuevos_columnas
    return df_limpio


def encontrar_columna_index(df, posibles_nombres):
    indices = [i for i, col in enumerate(df.columns) if any(p in str(col) for p in posibles_nombres)]
    return indices[-1] if indices else None


if archivo_bodega and archivo_farmacia and archivo_solicitud:
    if st.button("🚀 Procesar Despacho de Lotes", type="primary", use_container_width=True):
        try:
            with st.spinner("Procesando datos y aplicando reglas FEFO..."):
                df_bodega = optimizar_dataframe(archivo_bodega, ['COD', 'ARTICULO', 'LOTE'])
                df_inv_farmacia = optimizar_dataframe(archivo_farmacia, ['COD', 'ARTICULO', 'LOTE'])
                df_solicitud = optimizar_dataframe(archivo_solicitud, ['CÓD', 'COD', 'SOLICITADA', 'DESCRIP'])

                # Índices de columnas
                idx_cod_bodega = encontrar_columna_index(df_bodega, ['COD', 'ARTICULO'])
                idx_lote_bodega = encontrar_columna_index(df_bodega, ['NO LOTE', 'LOTE'])
                idx_venc_bodega = encontrar_columna_index(df_bodega, ['VENCIMIENTO', 'FECHA'])
                idx_cant_bodega_num = [i for i, c in enumerate(df_bodega.columns) if 'CANTIDAD' in str(c)][-1]

                idx_cod_farmacia = encontrar_columna_index(df_inv_farmacia, ['COD', 'ARTICULO'])
                idx_lote_farmacia = encontrar_columna_index(df_inv_farmacia, ['NO LOTE', 'LOTE'])
                idx_cant_farmacia_num = [i for i, c in enumerate(df_inv_farmacia.columns) if 'CANTIDAD' in str(c)][
                    -1] if 'CANTIDAD' in "".join(df_inv_farmacia.columns) else None

                idx_cod_solicitud = encontrar_columna_index(df_solicitud, ['CÓDIGO', 'CODIGO', 'COD'])
                idx_cant_solicitud = encontrar_columna_index(df_solicitud, ['SOLICITADA', 'CANTIDAD'])

                # Limpieza de filas vacías
                df_solicitud = df_solicitud.dropna(subset=[df_solicitud.columns[idx_cod_solicitud]])

                # Manejo de fechas y stock en Bodega
                col_venc_nombre = df_bodega.columns[idx_venc_bodega]
                df_bodega[col_venc_nombre] = pd.to_datetime(df_bodega[col_venc_nombre], dayfirst=True, errors='coerce')
                df_bodega = df_bodega.dropna(subset=[col_venc_nombre])

                col_cant_bod_nombre = df_bodega.columns[idx_cant_bodega_num]
                df_bodega[col_cant_bod_nombre] = pd.to_numeric(df_bodega[col_cant_bod_nombre], errors='coerce').fillna(
                    0)
                df_bodega = df_bodega[df_bodega[col_cant_bod_nombre] > 0]

                # Filtro de stock en Farmacia
                if idx_cant_farmacia_num is not None:
                    col_cant_farm_nombre = df_inv_farmacia.columns[idx_cant_farmacia_num]
                    df_inv_farmacia[col_cant_farm_nombre] = pd.to_numeric(df_inv_farmacia[col_cant_farm_nombre],
                                                                          errors='coerce').fillna(0)
                    df_inv_farmacia = df_inv_farmacia[df_inv_farmacia[col_cant_farm_nombre] > 0]

                # Ordenar por FEFO
                col_cod_bod_nombre = df_bodega.columns[idx_cod_bodega]
                df_bodega = df_bodega.sort_values(by=[col_cod_bod_nombre, col_venc_nombre]).reset_index(drop=True)
                df_bodega['PRIORIDAD_ORDEN'] = df_bodega.groupby(col_cod_bod_nombre).cumcount() + 1

                prioridades_texto = {1: "primero", 2: "segundo", 3: "tercero", 4: "cuarto"}
                lotes_sugeridos, cantidades_enviar, observaciones = [], [], []

                for idx, fila in df_solicitud.iterrows():
                    cod = fila.iloc[idx_cod_solicitud]
                    try:
                        cod = int(float(str(cod).strip()))
                    except:
                        lotes_sugeridos.append("Código Inválido")
                        cantidades_enviar.append(0)
                        observaciones.append("Revisar formato de código")
                        continue

                    cant_solicitada = fila.iloc[idx_cant_solicitud] if idx_cant_solicitud is not None else 0
                    if pd.isna(cant_solicitada):
                        cant_solicitada = 0

                    lote_actual_farmacia = ""
                    col_cod_farm_nombre = df_inv_farmacia.columns[idx_cod_farmacia]
                    col_lote_farm_nombre = df_inv_farmacia.columns[idx_lote_farmacia]

                    registro_farmacia = df_inv_farmacia[df_inv_farmacia[col_cod_farm_nombre] == cod]
                    if not registro_farmacia.empty:
                        lote_actual_farmacia = str(registro_farmacia.iloc[0][col_lote_farm_nombre]).strip()

                    lotes_en_bodega = df_bodega[df_bodega[col_cod_bod_nombre] == cod]

                    if lotes_en_bodega.empty:
                        lotes_sugeridos.append("Sin stock")
                        cantidades_enviar.append(0)
                        observaciones.append("No hay lotes disponibles en bodega")
                        continue

                    col_lote_bod_nombre = df_bodega.columns[idx_lote_bodega]
                    lote_fefo_row = lotes_en_bodega[lotes_en_bodega['PRIORIDAD_ORDEN'] == 1].iloc[0]
                    lote_a_enviar = lote_fefo_row[col_lote_bod_nombre]

                    lote_farmacia_en_bodega = lotes_en_bodega[
                        lotes_en_bodega[col_lote_bod_nombre] == lote_actual_farmacia]

                    if lote_actual_farmacia and not lote_farmacia_en_bodega.empty:
                        prioridad_farmacia = lote_farmacia_en_bodega.iloc[0]['PRIORIDAD_ORDEN']
                        texto_prio = prioridades_texto.get(prioridad_farmacia, f"{prioridad_farmacia}O")
                        if prioridad_farmacia == 1:
                            observacion = "mismo lote que esta saliendo"
                        else:
                            observacion = f"mismo lote que sale {texto_prio}"
                    else:
                        prioridad_sugerido = lote_fefo_row['PRIORIDAD_ORDEN']
                        texto_prio_sug = prioridades_texto.get(prioridad_sugerido, f"{prioridad_sugerido}O")
                        observacion = f"sale {texto_prio_sug}"

                    lotes_sugeridos.append(lote_a_enviar)
                    cantidades_enviar.append(cant_solicitada)
                    observaciones.append(observacion)

                # Guardar resultados
                idx_out_enviada = encontrar_columna_index(df_solicitud, ['ENVIADA']) or 'CANTIDAD ENVIADA'
                idx_out_lote = encontrar_columna_index(df_solicitud, ['LOTE']) or 'LOTE'

                df_solicitud[df_solicitud.columns[idx_out_enviada] if isinstance(idx_out_enviada,
                                                                                 int) else idx_out_enviada] = cantidades_enviar
                df_solicitud[df_solicitud.columns[idx_out_lote] if isinstance(idx_out_lote,
                                                                              int) else idx_out_lote] = lotes_sugeridos
                df_solicitud['OBSERVACIÓN'] = observaciones

                # Convertir a binario para descarga en la web
                output = io.BytesIO()
                with pd.ExcelWriter(output, engine='openpyxl') as writer:
                    df_solicitud.to_excel(writer, index=False)
                datos_excel = output.getvalue()

                st.success("🎉 ¡Análisis completado con éxito!")
                st.download_button(
                    label="📥 Descargar Reporte de Despacho Listo",
                    data=datos_excel,
                    file_name="Reporte_Despacho_Lotes.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
        except Exception as e:
            st.error(f"Ocurrió un problema técnico: {str(e)}")
else:
    st.info("💡 Por favor, carga los 3 archivos de Excel en la barra lateral izquierda para iniciar.")