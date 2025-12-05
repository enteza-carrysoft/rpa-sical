import os
import sys
import time
from datetime import datetime

import pandas as pd
import pyautogui
import pyperclip
import PySimpleGUI as sg
import openpyxl  # para que PyInstaller lo incluya si empaquetas


# ---------------------------------------------------------
# CONFIGURACIÓN GLOBAL
# ---------------------------------------------------------

pyautogui.FAILSAFE = True  # mover ratón a esquina sup. izda aborta

# OpenCV (opcional, para usar 'confidence' en locateOnScreen)
try:
    import cv2  # noqa: F401
    OPENCV_AVAILABLE = True
except ImportError:
    OPENCV_AVAILABLE = False


# ---------------------------------------------------------
# RUTAS DE RECURSOS (COMPATIBLE CON PYINSTALLER)
# ---------------------------------------------------------

def resource_path(relative_path: str) -> str:
    """
    Devuelve la ruta absoluta a un recurso, funcionando tanto en ejecución
    normal como en modo PyInstaller.
    """
    if hasattr(sys, "_MEIPASS"):
        base_path = sys._MEIPASS  # type: ignore[attr-defined]
    else:
        base_path = os.path.abspath(".")
    return os.path.join(base_path, relative_path)


OPERACION_IMG = resource_path("images/campo_operacion.png")
VALIDAR_IMG = resource_path("images/boton_validar.png")
YES_IMG = resource_path("images/boton_yes.png")
LOGO_IMG = resource_path("images/logo.png")  # logo Diputación

# Imágenes para mensajes de SICAL
MSG_SIMPLE_IMG = resource_path("images/msg_simple.png")        # ventana de aviso "normal"
MSG_BTN_ACEPTAR_IMG = resource_path("images/btn_aceptar.png")  # botón Aceptar
MSG_CRITICO_IMG = resource_path("images/msg_critico.png")      # ventana de error crítico


# ---------------------------------------------------------
# UTILIDADES DE ESCRITURA Y LOCALIZACIÓN
# ---------------------------------------------------------

def write_fast(text: str):
    """
    Escribe texto en el campo activo.
    - <= 20 caracteres: tecleo con pyautogui.write
    - > 20 caracteres: pega desde portapapeles (mejor con tildes, etc.)
    """
    if text is None:
        return

    text = str(text)
    if not text:
        return

    if len(text) > 20:
        pyperclip.copy(text)
        pyautogui.hotkey("ctrl", "v")
        time.sleep(0.05)
    else:
        pyautogui.write(text, interval=0.02)


def localizar_en_pantalla(ruta_imagen: str, confidence: float):
    """
    Busca una imagen en pantalla.
    - Si hay OpenCV, usa 'confidence'.
    - Si no, coincidencia exacta sin 'confidence'.
    Devuelve el centro (x,y) o None.
    """
    if not os.path.exists(ruta_imagen):
        return None

    if OPENCV_AVAILABLE:
        loc = pyautogui.locateCenterOnScreen(ruta_imagen, confidence=confidence)
    else:
        loc = pyautogui.locateCenterOnScreen(ruta_imagen)

    return loc


# ---------------------------------------------------------
# GESTIÓN DE MENSAJES DE SICAL / CONTEXTO
# ---------------------------------------------------------

def comprobar_mensajes_sical(window, confidence=0.8, delay_click=0.3) -> str:
    """
    Revisa si ha aparecido algún mensaje modal conocido de SICAL.

    Devuelve:
      - "NINGUNO"     -> no se ha detectado nada.
      - "AUTOCERRADO" -> mensaje simple detectado y cerrado (Aceptar).
      - "CRITICO"     -> mensaje crítico detectado (el RPA debe pararse).
    """

    # 1) Mensaje crítico
    if os.path.exists(MSG_CRITICO_IMG):
        loc_crit = localizar_en_pantalla(MSG_CRITICO_IMG, confidence)
        if loc_crit:
            window["-LOG-"].print("Mensaje CRÍTICO detectado en SICAL.")
            window["-STATUS-"].update("Mensaje crítico detectado. Deteniendo RPA.")
            return "CRITICO"

    # 2) Mensaje simple (aviso) autocerrable
    if os.path.exists(MSG_SIMPLE_IMG):
        loc_msg = localizar_en_pantalla(MSG_SIMPLE_IMG, confidence)
        if loc_msg:
            if os.path.exists(MSG_BTN_ACEPTAR_IMG):
                loc_btn = localizar_en_pantalla(MSG_BTN_ACEPTAR_IMG, confidence)
                if loc_btn:
                    pyautogui.click(loc_btn)
                    time.sleep(delay_click)
                    window["-LOG-"].print(
                        "Mensaje de aviso en SICAL detectado y cerrado automáticamente."
                    )
                    return "AUTOCERRADO"

    return "NINGUNO"


def esperar_campo_operacion(window, confidence: float, timeout: float = 5.0) -> bool:
    """
    Espera hasta 'timeout' segundos a que reaparezca el campo 'Operación'
    en pantalla.

    Devuelve True si lo encuentra, False si no.
    """
    inicio = time.time()
    while time.time() - inicio < timeout:
        loc = localizar_en_pantalla(OPERACION_IMG, confidence)
        if loc:
            return True

        # Mientras esperamos, por si aparece un mensaje crítico
        estado_msg = comprobar_mensajes_sical(window, confidence=confidence)
        if estado_msg == "CRITICO":
            return False

        time.sleep(0.3)

    window["-LOG-"].print(
        "Tras validar no ha reaparecido el campo de 'Operación'. "
        "Probable mensaje o estado inesperado en SICAL."
    )
    window["-STATUS-"].update(
        "No se ha podido volver al campo 'Operación'. Revisa SICAL."
    )
    return False


# ---------------------------------------------------------
# LECTURA Y NORMALIZACIÓN DEL EXCEL
# ---------------------------------------------------------

def leer_excel_rpa(file_path: str) -> pd.DataFrame:
    """
    Lee el Excel (xlsx) y devuelve un DataFrame:
    - Nombres de columnas normalizados.
    - Columnas que contengan 'fecha' -> texto dd/mm/yyyy.
    - Columna 'Salto' -> siempre texto.
    """
    df = pd.read_excel(file_path, engine="openpyxl")
    df.columns = [str(c).strip() for c in df.columns]

    # Columnas de fecha por nombre
    for col in df.columns:
        col_lower = col.lower()
        if "fecha" in col_lower:
            serie = pd.to_datetime(df[col], dayfirst=True, errors="coerce")
            df[col] = serie.dt.strftime("%d/%m/%Y").fillna("")

    # Columna 'Salto' como texto
    for posible in ["Salto", "salto", "SALTO"]:
        if posible in df.columns:
            df[posible] = df[posible].astype(str).replace("nan", "")

    return df


def guess_importe_col(df: pd.DataFrame) -> str | None:
    """
    Intenta localizar la columna de importe.
    """
    for c in df.columns:
        if "importe" in str(c).lower():
            return c
    return df.columns[-1] if len(df.columns) > 0 else None


def guess_operacion_col(df: pd.DataFrame) -> str | None:
    """
    Intenta localizar una columna que identifique la operación
    (solo para mostrar info en pantalla).
    """
    for c in df.columns:
        nombre = str(c).lower()
        if "oper" in nombre or "op." in nombre:
            return c
    return df.columns[0] if len(df.columns) > 0 else None


# ---------------------------------------------------------
# LÓGICA DEL RPA (SIN HILOS)
# ---------------------------------------------------------

def ejecutar_rpa(
    window,
    df: pd.DataFrame,
    importe_col: str,
    delay_tab: float,
    delay_click: float,
    confidence: float,
):
    """
    Ejecuta el RPA de forma síncrona (sin hilos).
    Se apoya en:
    - comprobar_mensajes_sical
    - esperar_campo_operacion
    - pyautogui.FAILSAFE (mover ratón a la esquina sup. izda aborta)
    """

    try:
        if importe_col not in df.columns:
            window["-LOG-"].print(
                f"ERROR: La columna de importe '{importe_col}' no existe en el Excel."
            )
            window["-STATUS-"].update(
                f"La columna importe '{importe_col}' no existe."
            )
            return

        oper_col = guess_operacion_col(df)

        window["-STATUS-"].update("RPA iniciado. Preparando entorno...")
        window["-LOG-"].print("RPA iniciado.")
        window.refresh()
        time.sleep(1)

        window["-STATUS-"].update(
            "Tienes 5 segundos para poner SICAL en primer plano, en el campo 'Operación'..."
        )
        window["-LOG-"].print(
            "Pon SICAL en primer plano, con el foco en el campo 'Operación'."
        )
        window.refresh()
        time.sleep(5)

        total = len(df)

        for idx, row in df.iterrows():
            # Info de operación e importe
            op_text = ""
            if oper_col and oper_col in df.columns:
                op_text = str(row[oper_col])

            valor_importe = row[importe_col]
            valor_importe_str = "" if pd.isna(valor_importe) else str(valor_importe).strip()

            window["-OP-"].update(f"{op_text} (fila {idx + 1}/{total})")
            window["-IMP-"].update(valor_importe_str)
            window["-STATUS-"].update(f"Procesando fila {idx + 1} de {total}")
            window["-LOG-"].print(f"Procesando fila {idx + 1} de {total}")
            window.refresh()

            # 1) Buscar campo operación
            window["-LOG-"].print("Buscando campo 'Operación' en pantalla...")
            window.refresh()
            loc_op = localizar_en_pantalla(OPERACION_IMG, confidence)
            if loc_op is None:
                window["-LOG-"].print(
                    f"ERROR: No se pudo localizar el campo de operación ({OPERACION_IMG})."
                )
                window["-STATUS-"].update(
                    "No se pudo localizar el campo 'Operación'. Se detiene el RPA."
                )
                return

            pyautogui.click(loc_op)
            time.sleep(delay_click)

            # Mensajes justo al hacer click
            estado_msg = comprobar_mensajes_sical(
                window, confidence=confidence, delay_click=delay_click
            )
            if estado_msg == "CRITICO":
                return

            # 2) Recorrer columnas del DataFrame
            for col in df.columns:
                valor = row[col]
                valor_str = "" if pd.isna(valor) else str(valor).strip()

                # 'T' => solo TAB
                if valor_str.upper() == "T":
                    pyautogui.press("tab")
                    time.sleep(delay_tab)

                    estado_msg = comprobar_mensajes_sical(
                        window, confidence=confidence, delay_click=delay_click
                    )
                    if estado_msg == "CRITICO":
                        return

                    if col == importe_col:
                        window["-LOG-"].print(
                            "Advertencia: la columna importe tiene 'T'; no se introduce importe."
                        )
                    continue

                # Si es importe, punto -> coma
                if col == importe_col:
                    valor_str = valor_str.replace(".", ",")

                write_fast(valor_str)
                pyautogui.press("tab")
                time.sleep(delay_tab)

                estado_msg = comprobar_mensajes_sical(
                    window, confidence=confidence, delay_click=delay_click
                )
                if estado_msg == "CRITICO":
                    return

                # Al llegar al importe: Validar + Sí
                if col == importe_col:
                    window["-LOG-"].print("Buscando botón 'Validar'...")
                    window.refresh()
                    loc_val = localizar_en_pantalla(VALIDAR_IMG, confidence)
                    if loc_val is None:
                        window["-LOG-"].print(
                            f"ERROR: No se encontró el botón 'Validar' ({VALIDAR_IMG})."
                        )
                        window["-STATUS-"].update(
                            "No se encontró botón 'Validar'. Se detiene el RPA."
                        )
                        return
                    pyautogui.click(loc_val)
                    time.sleep(delay_click)

                    estado_msg = comprobar_mensajes_sical(
                        window, confidence=confidence, delay_click=delay_click
                    )
                    if estado_msg == "CRITICO":
                        return

                    window["-LOG-"].print("Buscando botón 'Sí'...")
                    window.refresh()
                    loc_yes = localizar_en_pantalla(YES_IMG, confidence)
                    if loc_yes is None:
                        window["-LOG-"].print(
                            f"ERROR: No se encontró el botón 'Sí' ({YES_IMG})."
                        )
                        window["-STATUS-"].update(
                            "No se encontró botón 'Sí'. Se detiene el RPA."
                        )
                        return
                    pyautogui.click(loc_yes)
                    time.sleep(delay_click)

                    estado_msg = comprobar_mensajes_sical(
                        window, confidence=confidence, delay_click=delay_click
                    )
                    if estado_msg == "CRITICO":
                        return

                    # Comprobar que volvemos al campo Operación
                    if not esperar_campo_operacion(window, confidence=confidence, timeout=5.0):
                        return

                    window["-LOG-"].print(
                        f"Registro de la fila {idx + 1} confirmado correctamente."
                    )
                    window.refresh()
                    break

        window["-STATUS-"].update("RPA finalizado. Todas las filas procesadas.")
        window["-LOG-"].print("RPA finalizado correctamente.")
        window.refresh()

    except pyautogui.FailSafeException:
        window["-LOG-"].print("RPA abortado por FAILSAFE (ratón a esquina sup. izda).")
        window["-STATUS-"].update("RPA abortado por FAILSAFE.")
    except Exception as e:
        window["-LOG-"].print(f"ERROR inesperado en el RPA: {e}")
        window["-STATUS-"].update(f"ERROR inesperado: {e}")


# ---------------------------------------------------------
# INTERFAZ PySimpleGUI
# ---------------------------------------------------------

def crear_ventana():
    sg.theme("SystemDefaultForReal")

    # Logo + título
    top_row = []
    if os.path.exists(LOGO_IMG):
        top_row.append(sg.Image(filename=LOGO_IMG, size=(200, 80), pad=((0, 20), (0, 0))))
    top_row.append(
        sg.Column(
            [
                [sg.Text("RPA Operaciones de Gasto en SICAL", font=("Segoe UI", 18, "bold"))],
                [sg.Text("Diputación Provincial de Sevilla", font=("Segoe UI", 11))],
                [sg.Text(
                    "Consejo: si algo va mal, mueve el ratón a la esquina superior izquierda\n"
                    "para abortar inmediatamente (pyautogui FAILSAFE).",
                    font=("Segoe UI", 8),
                    text_color="#555555"
                )]
            ],
            vertical_alignment="top"
        )
    )

    # Sección archivo
    frame_archivo = sg.Frame(
        "1. Archivo de operaciones (Excel .xlsx)",
        [
            [
                sg.Input(key="-FILE-", expand_x=True, disabled=True),
                sg.FileBrowse("Buscar...", file_types=(("Excel", "*.xlsx"),)),
                sg.Button("Cargar Excel", key="-LOAD-")
            ],
            [sg.Text("Vista previa (primeras filas):")],
            [
                sg.Table(
                    values=[],
                    headings=[],
                    key="-TABLE-",
                    auto_size_columns=True,
                    num_rows=8,
                    justification="left",
                    enable_events=False,
                    expand_x=True,
                    expand_y=False,
                    alternating_row_color="#F7F7F7",
                    row_height=18,
                )
            ],
        ],
        expand_x=True,
    )

    # Sección importe
    frame_importe = sg.Frame(
        "2. Configuración de importe",
        [
            [
                sg.Text("Columna de importe:"),
                sg.Combo(
                    values=[],
                    key="-IMP_COL-",
                    size=(40, 1),
                    readonly=True
                )
            ]
        ],
        expand_x=True,
    )

    # Sección parámetros
    frame_param = sg.Frame(
        "3. Parámetros del RPA",
        [
            [
                sg.Text("Confianza detección imágenes:"),
                sg.Slider(
                    range=(50, 99),
                    default_value=80,
                    resolution=1,
                    orientation="h",
                    key="-CONF-",
                    size=(30, 15),
                ),
                sg.Text("(se ignora si no hay OpenCV)", font=("Segoe UI", 8))
            ],
            [
                sg.Text("Espera tras cada TAB (segundos):"),
                sg.Input("0.40", key="-DELAY_TAB-", size=(5, 1)),
                sg.Text("Espera tras cada click (segundos):"),
                sg.Input("0.60", key="-DELAY_CLICK-", size=(5, 1)),
            ],
        ],
        expand_x=True,
    )

    # Sección ejecución
    frame_ejec = sg.Frame(
        "4. Ejecución del RPA",
        [
            [sg.Text("Estado:", size=(8, 1)), sg.Text("", key="-STATUS-", size=(70, 1))],
            [
                sg.Text("Operación:", size=(8, 1)),
                sg.Text("", key="-OP-", size=(40, 1)),
                sg.Text("Importe:", size=(8, 1)),
                sg.Text("", key="-IMP-", size=(20, 1)),
            ],
            [
                sg.Button("▶ Iniciar RPA", key="-START-", size=(15, 1),
                          button_color=("white", "#00704A")),
                sg.Push(),
                sg.Button("Salir", key="-EXIT-")
            ],
        ],
        expand_x=True,
    )

    # Log
    log_frame = sg.Frame(
        "Registro de actividad",
        [
            [
                sg.Multiline(
                    key="-LOG-",
                    size=(100, 15),
                    autoscroll=True,
                    disabled=True,
                    font=("Consolas", 9),
                    background_color="#111111",
                    text_color="#EEEEEE",
                )
            ]
        ],
        expand_x=True,
        expand_y=True,
    )

    layout = [
        [*top_row],
        [frame_archivo],
        [frame_importe],
        [frame_param],
        [frame_ejec],
        [log_frame],
    ]

    return sg.Window(
        "RPA Operaciones de Gasto - Diputación de Sevilla",
        layout,
        resizable=True,
        finalize=True,
    )


# ---------------------------------------------------------
# BUCLE PRINCIPAL
# ---------------------------------------------------------

def main():
    window = crear_ventana()
    df = None

    while True:
        event, values = window.read()

        if event in (sg.WIN_CLOSED, "-EXIT-"):
            break

        # Cargar Excel
        if event == "-LOAD-":
            file_path = values["-FILE-"]
            if not file_path:
                sg.popup_error("Selecciona un fichero Excel (.xlsx).")
                continue

            if not os.path.exists(file_path):
                sg.popup_error("El fichero seleccionado no existe.")
                continue

            try:
                df = leer_excel_rpa(file_path)
            except Exception as e:
                sg.popup_error(f"No se pudo leer el Excel:\n{e}")
                continue

            # Actualizamos tabla de vista previa
            preview = df.head(8)
            data = preview.values.tolist()
            headings = list(preview.columns)
            window["-TABLE-"].update(values=data, headings=headings)

            # Rellenamos combo de importe
            imp_col = guess_importe_col(df)
            window["-IMP_COL-"].update(values=list(df.columns), value=imp_col)

            window["-STATUS-"].update("Excel cargado correctamente.")
            window["-LOG-"].update("")
            window["-LOG-"].print("Excel cargado correctamente.")

        # Iniciar RPA (bloqueante)
        if event == "-START-":
            if df is None:
                sg.popup_error("Primero debes cargar un archivo Excel.")
                continue

            importe_col = values["-IMP_COL-"]
            if not importe_col:
                sg.popup_error("Selecciona la columna de importe.")
                continue

            # Delays
            try:
                delay_tab = float(values["-DELAY_TAB-"].replace(",", "."))
                delay_click = float(values["-DELAY_CLICK-"].replace(",", "."))
            except ValueError:
                sg.popup_error("Los tiempos de espera deben ser números (ej. 0.4).")
                continue

            confidence = float(values["-CONF-"]) / 100.0  # slider 50-99 -> 0.50-0.99

            # Comprobamos que imágenes existan
            faltan = [
                p for p in [OPERACION_IMG, VALIDAR_IMG, YES_IMG]
                if not os.path.exists(p)
            ]
            if faltan:
                sg.popup_error(
                    "Faltan archivos de imagen para el RPA:\n" + "\n".join(faltan)
                )
                continue

            # Ejecutamos RPA (bloquea la ventana mientras corre)
            ejecutar_rpa(
                window=window,
                df=df,
                importe_col=importe_col,
                delay_tab=delay_tab,
                delay_click=delay_click,
                confidence=confidence,
            )

    window.close()


if __name__ == "__main__":
    main()
