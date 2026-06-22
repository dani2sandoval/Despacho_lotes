import streamlit as st
import pandas as pd
import io

st.set_page_config(page_title="Control de Rotación IGSS", page_icon="📦", layout="centered")

st.title("📦 Control de Rotación de Inventario")
st.subheader("Hospital General IGSS Jalapa")

archivo_bodega = st.sidebar.file_uploader("1. Inventario de Bodega (.xlsx)", type=["xlsx"])
archivo_farmacia = st.sidebar.file_uploader("2. Inventario de Farmacia (.xlsx)", type=["xlsx"])
archivo_solicitud = st.sidebar.file_uploader("3. Solicitud de Farmacia (.xlsx)", type=["xlsx"])


def leer_excel_limpio(archivo, palabras_clave):
    df_sucio = pd.read_excel(archivo, header=None)

    fila_encabezado = 0
    for i, fila in df_sucio.iterrows():
        valores = [str(x).strip().upper() for x in fila.dropna()]
        if any(any(p in valor for p in palabras_clave) for valor in valores):
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


def texto_orden(n):
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
    return textos.get(n, f"{n}°")


def preparar_inventario(df):
    col_codigo = buscar_columna(df, ["COD ARTICULO", "CÓD ARTICULO", "CODIGO", "CÓDIGO", "COD"])
    col_medicamento = buscar_columna(df, ["MEDICAMENTO", "DESCRIPCION", "DESCRIPCIÓN"])
    col_lote = buscar_columna(df, ["NO LOTE", "LOTE"])
    col_vencimiento = buscar_columna(df, ["FECHA VENCIMIENTO", "VENCIMIENTO", "FECHA"])
    col_cantidad = buscar_columna(df, ["CANTIDAD", "EXISTENCIA"])

    faltantes = []
    if col_codigo is None:
        faltantes.append("Código")
    if col_lote is None:
        faltantes.append("Lote")
    if col_vencimiento is None:
        faltantes.append("Fecha de vencimiento")
    if col_cantidad is None:
        faltantes.append("Cantidad")

    if faltantes:
        raise Exception("Faltan columnas: " + ", ".join(faltantes))

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


def calcular_prioridad_real(lote_enviar, fecha_enviar, bodega, farmacia):
    comparacion = []

    for _, fila in bodega.iterrows():
        comparacion.append({
            "LOTE": fila["LOTE_LIMPIO"],
            "FECHA": fila["VENCIMIENTO_LIMPIO"],
            "CANTIDAD": fila["CANTIDAD_LIMPIA"],
            "ORIGEN": "BODEGA"
        })

    for _, fila in farmacia.iterrows():
        comparacion.append({
            "LOTE": fila["LOTE_LIMPIO"],
            "FECHA": fila["VENCIMIENTO_LIMPIO"],
            "CANTIDAD": fila["CANTIDAD_LIMPIA"],
            "ORIGEN": "FARMACIA"
        })

    df_comp = pd.DataFrame(comparacion)

    df_comp = df_comp.sort_values(
        by=["FECHA", "CANTIDAD"],
        ascending=[True, True]
    ).reset_index(drop=True)

    df_comp["PRIORIDAD_REAL"] = df_comp.index + 1

    fila_lote = df_comp[
        (df_comp["LOTE"] == lote_enviar) &
        (df_comp["FECHA"] == fecha_enviar) &
        (df_comp["ORIGEN"] == "BODEGA")
    ]

    if fila_lote.empty:
        return 1

    return int(fila_lote.iloc[0]["PRIORIDAD_REAL"])


def seleccionar_lote(codigo, df_bodega, df_farmacia):
    bodega = df_bodega[df_bodega["CODIGO_LIMPIO"] == codigo].copy()
    farmacia = df_farmacia[df_farmacia["CODIGO_LIMPIO"] == codigo].copy()

    if bodega.empty:
        return "", 0, "sin stock en bodega"

    bodega = bodega.sort_values(
        by=["VENCIMIENTO_LIMPIO", "CANTIDAD_LIMPIA"],
        ascending=[True, True]
    ).reset_index(drop=True)

    primera_fecha_bodega = bodega["VENCIMIENTO_LIMPIO"].min()

    grupo_primera_fecha_bodega = bodega[
        bodega["VENCIMIENTO_LIMPIO"] == primera_fecha_bodega
    ].copy()

    lotes_farmacia = set(farmacia["LOTE_LIMPIO"].tolist())

    coincidencia_misma_fecha_bodega = grupo_primera_fecha_bodega[
        grupo_primera_fecha_bodega["LOTE_LIMPIO"].isin(lotes_farmacia)
    ]

    if not coincidencia_misma_fecha_bodega.empty:
        lote_elegido = coincidencia_misma_fecha_bodega.iloc[0]
    else:
        lote_elegido = grupo_primera_fecha_bodega.sort_values(
            by="CANTIDAD_LIMPIA",
            ascending=True
        ).iloc[0]

    lote_enviar = lote_elegido["LOTE_LIMPIO"]
    fecha_enviar = lote_elegido["VENCIMIENTO_LIMPIO"]
    cantidad_lote_bodega = lote_elegido["CANTIDAD_LIMPIA"]

    if farmacia.empty:
        prioridad_real = calcular_prioridad_real(lote_enviar, fecha_enviar, bodega, farmacia)
        return lote_enviar, cantidad_lote_bodega, f"sale {texto_orden(prioridad_real)}"

    farmacia = farmacia.sort_values(
        by=["VENCIMIENTO_LIMPIO", "CANTIDAD_LIMPIA"],
        ascending=[True, True]
    ).reset_index(drop=True)

    lote_farmacia = farmacia.iloc[0]["LOTE_LIMPIO"]
    fecha_farmacia = farmacia.iloc[0]["VENCIMIENTO_LIMPIO"]

    if lote_farmacia == lote_enviar:
        observacion = "mismo lote"

    elif fecha_farmacia == fecha_enviar:
        observacion = "misma fecha de vencimiento del lote que está saliendo"

    else:
        prioridad_real = calcular_prioridad_real(lote_enviar, fecha_enviar, bodega, farmacia)
        observacion = f"sale {texto_orden(prioridad_real)}"

    return lote_enviar, cantidad_lote_bodega, observacion


if archivo_bodega and archivo_farmacia and archivo_solicitud:
    if st.button("🚀 Procesar Despacho", type="primary", use_container_width=True):
        try:
            with st.spinner("Procesando datos..."):

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

                lotes = []
                cantidades = []
                observaciones = []
                existencias = []

                for _, fila in df_solicitud.iterrows():
                    codigo = fila["CODIGO_LIMPIO"]
                    cantidad_solicitada = fila["CANTIDAD_SOLICITADA_LIMPIA"]

                    lote, existencia, observacion = seleccionar_lote(
                        codigo,
                        df_bodega,
                        df_farmacia
                    )

                    lotes.append(lote)
                    cantidades.append(cantidad_solicitada)
                    existencias.append(existencia)
                    observaciones.append(observacion)

                resultado = df_solicitud.copy()
                resultado["CANTIDAD ENVIADA"] = cantidades
                resultado["LOTE"] = lotes
                resultado["EXISTENCIA LOTE BODEGA"] = existencias
                resultado["OBSERVACIÓN"] = observaciones

                resultado = resultado.drop(
                    columns=["CODIGO_LIMPIO", "CANTIDAD_SOLICITADA_LIMPIA"],
                    errors="ignore"
                )

                output = io.BytesIO()

                with pd.ExcelWriter(output, engine="openpyxl") as writer:
                    resultado.to_excel(writer, index=False, sheet_name="Despacho")

                st.success("✅ Reporte generado correctamente.")

                st.download_button(
                    label="📥 Descargar Reporte",
                    data=output.getvalue(),
                    file_name="Reporte_Despacho_Lotes.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )

                st.dataframe(resultado, use_container_width=True)

        except Exception as e:
            st.error(f"Ocurrió un error: {str(e)}")

else:
    st.info("Carga los 3 archivos Excel para iniciar.")
