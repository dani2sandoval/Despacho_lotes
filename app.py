import streamlit as st
import pandas as pd
import io

st.set_page_config(
    page_title="Control de Rotación IGSS",
    page_icon="📦",
    layout="centered"
)

st.title("📦 Control de Rotación de Inventario FEFO")
st.subheader("Hospital General IGSS Jalapa")
st.markdown("Carga los archivos de bodega, farmacia y solicitud para generar el despacho correcto de lotes.")

st.sidebar.header("📁 Carga de Documentos")

archivo_bodega = st.sidebar.file_uploader("1. Inventario de Bodega (.xlsx)", type=["xlsx"])
archivo_farmacia = st.sidebar.file_uploader("2. Existencia de Farmacia (.xlsx)", type=["xlsx"])
archivo_solicitud = st.sidebar.file_uploader("3. Solicitud de Farmacia (.xlsx)", type=["xlsx"])


def leer_excel_limpio(archivo, palabras_clave):
    df_sucio = pd.read_excel(archivo, header=None)

    fila_encabezado = 0
    for i, fila in df_sucio.iterrows():
        valores = [str(x).strip().upper() for x in fila.dropna()]
        if any(any(palabra in valor for palabra in palabras_clave) for valor in valores):
            fila_encabezado = i
            break

    df = pd.read_excel(archivo, skiprows=fila_encabezado)
    df.columns = [str(c).strip().upper().replace("\n", " ") for c in df.columns]
    return df


def buscar_columna(df, opciones):
    for opcion in opciones:
        for col in df.columns:
            if opcion in str(col).upper():
                return col
    return None


def normalizar_codigo(valor):
    if pd.isna(valor):
        return ""
    try:
        return str(int(float(str(valor).strip())))
    except:
        return str(valor).strip().upper()


def normalizar_lote(valor):
    if pd.isna(valor):
        return ""
    return str(valor).strip().upper()


def texto_orden(numero):
    textos = {
        1: "primero",
        2: "segundo",
        3: "tercero",
        4: "cuarto",
        5: "quinto",
        6: "sexto",
        7: "séptimo",
        8: "octavo",
        9: "noveno",
        10: "décimo"
    }
    return textos.get(numero, f"{numero}°")


def preparar_inventario(df):
    col_codigo = buscar_columna(df, ["COD ARTICULO", "CÓD ARTICULO", "CODIGO", "CÓDIGO", "COD"])
    col_medicamento = buscar_columna(df, ["MEDICAMENTO", "DESCRIPCION", "DESCRIPCIÓN"])
    col_lote = buscar_columna(df, ["NO LOTE", "LOTE"])
    col_vencimiento = buscar_columna(df, ["FECHA VENCIMIENTO", "VENCIMIENTO", "FECHA"])
    col_cantidad = buscar_columna(df, ["CANTIDAD", "EXISTENCIA"])

    columnas_faltantes = []
    if col_codigo is None:
        columnas_faltantes.append("Código")
    if col_lote is None:
        columnas_faltantes.append("Lote")
    if col_vencimiento is None:
        columnas_faltantes.append("Fecha vencimiento")
    if col_cantidad is None:
        columnas_faltantes.append("Cantidad")

    if columnas_faltantes:
        raise Exception(f"Faltan columnas en inventario: {', '.join(columnas_faltantes)}")

    df = df.copy()

    df["CODIGO_LIMPIO"] = df[col_codigo].apply(normalizar_codigo)
    df["LOTE_LIMPIO"] = df[col_lote].apply(normalizar_lote)
    df["VENCIMIENTO_LIMPIO"] = pd.to_datetime(df[col_vencimiento], dayfirst=True, errors="coerce")
    df["CANTIDAD_LIMPIA"] = pd.to_numeric(df[col_cantidad], errors="coerce").fillna(0)

    if col_medicamento:
        df["MEDICAMENTO_LIMPIO"] = df[col_medicamento].astype(str).str.strip()
    else:
        df["MEDICAMENTO_LIMPIO"] = ""

    df = df[
        (df["CODIGO_LIMPIO"] != "") &
        (df["LOTE_LIMPIO"] != "") &
        (df["VENCIMIENTO_LIMPIO"].notna()) &
        (df["CANTIDAD_LIMPIA"] > 0)
    ].copy()

    return df


def preparar_solicitud(df):
    col_codigo = buscar_columna(df, ["COD ARTICULO", "CÓD ARTICULO", "CODIGO", "CÓDIGO", "COD"])
    col_cantidad = buscar_columna(df, ["SOLICITADA", "CANTIDAD"])

    if col_codigo is None:
        raise Exception("No se encontró la columna de código en la solicitud.")

    df = df.copy()
    df["CODIGO_LIMPIO"] = df[col_codigo].apply(normalizar_codigo)

    if col_cantidad:
        df["CANTIDAD_SOLICITADA_LIMPIA"] = pd.to_numeric(df[col_cantidad], errors="coerce").fillna(0)
    else:
        df["CANTIDAD_SOLICITADA_LIMPIA"] = 0

    df = df[df["CODIGO_LIMPIO"] != ""].copy()

    return df


def seleccionar_lote(codigo, df_bodega, df_farmacia):
    lotes_bodega = df_bodega[df_bodega["CODIGO_LIMPIO"] == codigo].copy()

    if lotes_bodega.empty:
        return "", 0, "Sin stock en bodega"

    # Primero se busca la fecha más próxima
    fecha_mas_proxima = lotes_bodega["VENCIMIENTO_LIMPIO"].min()

    grupo_fecha = lotes_bodega[
        lotes_bodega["VENCIMIENTO_LIMPIO"] == fecha_mas_proxima
    ].copy()

    lotes_farmacia = df_farmacia[df_farmacia["CODIGO_LIMPIO"] == codigo].copy()
    lotes_farmacia_set = set(lotes_farmacia["LOTE_LIMPIO"].tolist())

    # Regla especial:
    # Si farmacia tiene un lote que vence en la fecha más próxima, enviar ese mismo lote.
    coincidencias_misma_fecha = grupo_fecha[
        grupo_fecha["LOTE_LIMPIO"].isin(lotes_farmacia_set)
    ].copy()

    if not coincidencias_misma_fecha.empty:
        lote_elegido = coincidencias_misma_fecha.sort_values(
            by=["CANTIDAD_LIMPIA"],
            ascending=True
        ).iloc[0]

        return (
            lote_elegido["LOTE_LIMPIO"],
            lote_elegido["CANTIDAD_LIMPIA"],
            "mismo lote que está saliendo"
        )

    # Si farmacia no tiene ningún lote de esa fecha, se usa menor cantidad
    lote_elegido = grupo_fecha.sort_values(
        by=["CANTIDAD_LIMPIA"],
        ascending=True
    ).iloc[0]

    return (
        lote_elegido["LOTE_LIMPIO"],
        lote_elegido["CANTIDAD_LIMPIA"],
        "sale primero"
    )


if archivo_bodega and archivo_farmacia and archivo_solicitud:
    if st.button("🚀 Procesar Despacho de Lotes", type="primary", use_container_width=True):
        try:
            with st.spinner("Procesando inventarios y aplicando reglas de despacho..."):

                df_bodega_raw = leer_excel_limpio(
                    archivo_bodega,
                    ["COD", "ARTICULO", "MEDICAMENTO", "LOTE"]
                )

                df_farmacia_raw = leer_excel_limpio(
                    archivo_farmacia,
                    ["COD", "ARTICULO", "MEDICAMENTO", "LOTE"]
                )

                df_solicitud_raw = leer_excel_limpio(
                    archivo_solicitud,
                    ["COD", "CÓD", "SOLICITADA", "DESCRIP"]
                )

                df_bodega = preparar_inventario(df_bodega_raw)
                df_farmacia = preparar_inventario(df_farmacia_raw)
                df_solicitud = preparar_solicitud(df_solicitud_raw)

                lotes_enviar = []
                cantidades_enviar = []
                observaciones = []
                existencias_lote = []

                for _, fila in df_solicitud.iterrows():
                    codigo = fila["CODIGO_LIMPIO"]
                    cantidad_solicitada = fila["CANTIDAD_SOLICITADA_LIMPIA"]

                    lote, existencia_lote, observacion = seleccionar_lote(
                        codigo,
                        df_bodega,
                        df_farmacia
                    )

                    lotes_enviar.append(lote)
                    cantidades_enviar.append(cantidad_solicitada)
                    existencias_lote.append(existencia_lote)
                    observaciones.append(observacion)

                df_resultado = df_solicitud.copy()
                df_resultado["CANTIDAD ENVIADA"] = cantidades_enviar
                df_resultado["LOTE A ENVIAR"] = lotes_enviar
                df_resultado["EXISTENCIA LOTE BODEGA"] = existencias_lote
                df_resultado["OBSERVACIÓN"] = observaciones

                columnas_ocultar = [
                    "CODIGO_LIMPIO",
                    "CANTIDAD_SOLICITADA_LIMPIA"
                ]

                df_resultado = df_resultado.drop(
                    columns=[c for c in columnas_ocultar if c in df_resultado.columns],
                    errors="ignore"
                )

                output = io.BytesIO()
                with pd.ExcelWriter(output, engine="openpyxl") as writer:
                    df_resultado.to_excel(writer, index=False, sheet_name="Despacho")

                st.success("🎉 Análisis completado correctamente.")

                st.download_button(
                    label="📥 Descargar Reporte de Despacho",
                    data=output.getvalue(),
                    file_name="Reporte_Despacho_Lotes.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )

                st.dataframe(df_resultado, use_container_width=True)

        except Exception as e:
            st.error(f"Ocurrió un problema: {str(e)}")

else:
    st.info("💡 Carga los 3 archivos Excel en la barra lateral para iniciar.")
