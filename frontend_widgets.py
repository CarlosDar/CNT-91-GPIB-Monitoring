import tkinter as tk
from tkinter import ttk


def crear_layout_principal(root):
    """
    Crea la estructura principal de la interfaz gráfica, imitando el diseño de la imagen proporcionada, pero con calidad visual mejorada:
    - Panel lateral azul oscuro con botones y secciones
    - Barra superior blanca con estado de conexión
    - Área principal de contenido en gris claro
    - Botones con efecto hover y esquinas redondeadas (simuladas)
    - Entrada de texto para el identificador del dispositivo
    """
    # Colores
    azul_lateral = '#25364a'
    azul_boton = '#2980f2'
    azul_hover = '#3a8bfd'
    gris_fondo = '#f6f7fa'
    blanco = '#ffffff'
    gris_texto = '#6c757d'
    rojo_estado = '#e74c3c'
    azul_activo = '#1c2a3a'

    # Configuración de la ventana principal
    root.configure(bg=gris_fondo)
    root.geometry('1200x600')  # Resolución inicial más pequeña
    root.title('CNT-91')
    root.minsize(800, 400)  # Tamaño mínimo razonable
    root.resizable(True, True)  # Permitir redimensionar libremente

    style = ttk.Style()
    style.theme_use('clam')
    # Estilo para los botones laterales
    style.configure('Sidebar.TButton',
                    font=('Segoe UI', 12),
                    foreground='white',
                    background=azul_lateral,
                    borderwidth=0,
                    focusthickness=3,
                    focuscolor=azul_boton,
                    padding=10)
    style.map('Sidebar.TButton',
              background=[('active', azul_hover), ('!active', azul_lateral)],
              relief=[('pressed', 'sunken'), ('!pressed', 'flat')])
    # Estilo para el botón principal (Conectar)
    style.configure('PrimarySidebar.TButton',
                    font=('Segoe UI', 12, 'bold'),
                    foreground='white',
                    background=azul_boton,
                    borderwidth=0,
                    focusthickness=3,
                    focuscolor=azul_boton,
                    padding=10)
    style.map('PrimarySidebar.TButton',
              background=[('active', azul_hover), ('!active', azul_boton)],
              relief=[('pressed', 'sunken'), ('!pressed', 'flat')])
    
    # Estilo para el botón de desconectar (rojo)
    style.configure('DangerSidebar.TButton',
                    font=('Segoe UI', 12, 'bold'),
                    foreground='white',
                    background='#e74c3c',
                    borderwidth=0,
                    focusthickness=3,
                    focuscolor='#c0392b',
                    padding=10)
    style.map('DangerSidebar.TButton',
              background=[('active', '#c0392b'), ('!active', '#e74c3c')],
              relief=[('pressed', 'sunken'), ('!pressed', 'flat')])

    # Panel lateral
    frame_lateral = tk.Frame(root, bg=azul_lateral, width=220)
    frame_lateral.pack(side='left', fill='y')

    # Logo y título
    logo = tk.Label(frame_lateral, text='⚙️', bg=azul_lateral, fg='white', font=('Segoe UI', 22, 'bold'))
    logo.pack(pady=(24, 0))
    titulo = tk.Label(frame_lateral, text='CNT-91', bg=azul_lateral, fg='white', font=('Segoe UI', 18, 'bold'))
    titulo.pack(pady=(0, 10))

    # Entrada de texto para el identificador del dispositivo
    label_id = tk.Label(frame_lateral, text='Identificador GPIB:', bg=azul_lateral, fg='#b0b8c1', font=('Segoe UI', 10, 'bold'))
    label_id.pack(padx=20, anchor='w')
    entry_id = ttk.Entry(frame_lateral, font=('Segoe UI', 11))
    entry_id.insert(0, 'GPIB0::10::INSTR')
    entry_id.pack(fill='x', padx=20, pady=(0, 18))

    # Botón Conectar
    btn_conectar = ttk.Button(frame_lateral, text='🔌  Conectar', style='PrimarySidebar.TButton')
    btn_conectar.pack(fill='x', padx=20, pady=(0, 10))
    
    # Variable para rastrear el estado del botón (conectado/desconectado)
    btn_conectar.estado_conectado = False

    # Separador sutil
    sep1 = tk.Frame(frame_lateral, bg='#34495e', height=2)
    sep1.pack(fill='x', padx=20, pady=(0, 10))

    # Botón Allan Deviation vs tau (antes: Configuración)
    btn_config = ttk.Button(frame_lateral, text='Allan Deviation vs tau', style='Sidebar.TButton')
    btn_config.pack(fill='x', padx=20, pady=(0, 10))

    # Botón Frequency Datalogger
    btn_mediciones = ttk.Button(frame_lateral, text='📈  Frequency Datalogger', style='Sidebar.TButton')
    btn_mediciones.pack(fill='x', padx=20, pady=(0, 10))

    # Botón Información (icono profesional)
    btn_info = ttk.Button(frame_lateral, text='🛈  Información CNT-91', style='Sidebar.TButton')
    btn_info.pack(fill='x', padx=20, pady=(0, 10))

    # Simular esquinas redondeadas y efecto hover
    def on_enter(e):
        if e.widget == btn_conectar:
            if hasattr(e.widget, 'estado_conectado') and e.widget.estado_conectado:
                # Si está conectado, mantener estilo rojo
                e.widget.configure(style='DangerSidebar.TButton')
            else:
                # Si está desconectado, usar estilo azul
                e.widget.configure(style='PrimarySidebar.TButton')
        else:
            e.widget.configure(style='Sidebar.TButton')
        e.widget['cursor'] = 'hand2'
    
    def on_leave(e):
        if e.widget == btn_conectar:
            if hasattr(e.widget, 'estado_conectado') and e.widget.estado_conectado:
                # Si está conectado, mantener estilo rojo
                e.widget.configure(style='DangerSidebar.TButton')
            else:
                # Si está desconectado, usar estilo azul
                e.widget.configure(style='PrimarySidebar.TButton')
        else:
            e.widget.configure(style='Sidebar.TButton')
        e.widget['cursor'] = ''
    
    for btn in [btn_conectar, btn_config, btn_mediciones, btn_info]:
        btn.bind('<Enter>', on_enter)
        btn.bind('<Leave>', on_leave)

    # Barra superior
    frame_superior = tk.Frame(root, bg=blanco, height=50, highlightbackground='#e0e0e0', highlightthickness=1)
    frame_superior.pack(side='top', fill='x')
    frame_superior.pack_propagate(False)

    # Estado de conexión
    estado = tk.Label(frame_superior, text='Estado:  Desconectado', bg=blanco, fg=rojo_estado, font=('Segoe UI', 12, 'bold'))
    estado.pack(side='left', padx=30, pady=10)

    # Área principal de contenido
    frame_contenido = tk.Frame(root, bg=gris_fondo)
    frame_contenido.pack(expand=True, fill='both')

    # Retornar referencias útiles para enlazar lógica después
    return {
        'frame_lateral': frame_lateral,
        'entry_id': entry_id,
        'btn_conectar': btn_conectar,
        'btn_config': btn_config,
        'btn_mediciones': btn_mediciones,
        'btn_info': btn_info,
        'frame_superior': frame_superior,
        'estado': estado,
        'frame_contenido': frame_contenido
    } 

def get_info_cnt91_sections():
    """
    Devuelve la lista de secciones (título, texto) para la información del CNT-91.
    """
    return [
        ("Descripción General", """
El CNT-91 de Pendulum Instruments es un contador/temporizador/análisis de modulación de ultra alta resolución y rendimiento, diseñado para cubrir todas las necesidades de medición de tiempo y frecuencia en laboratorios de I+D, producción, control de calidad y calibración.

Está construido sobre una arquitectura basada en un oscilador de referencia interno de alta estabilidad, con opción de un oscilador de rubidio en la versión CNT-91R, que proporciona una excepcional estabilidad a largo plazo (típicamente mejor que 5×10⁻¹¹/día). En ausencia de esta opción, se puede seleccionar una base de tiempo OCXO de precisión o incluso operar con una referencia externa (10 MHz).

Internamente, el equipo utiliza técnicas de conteo recíproco interpolado, lo que le permite lograr resoluciones de hasta 12 dígitos por segundo de medición o una resolución temporal de 50 ps en modo de intervalo de tiempo. Este principio consiste en sincronizar la medición con los flancos del evento de entrada y luego aplicar interpolación analógica mediante carga de capacitores y posterior conversión ADC, reduciendo el error de cuantización a valores extremadamente bajos.

Su capacidad de medición "zero dead-time" lo hace único en su categoría, permitiendo capturar secuencias continuas de datos sin pérdida de información entre muestras, incluso en aplicaciones exigentes como ADEV, TIE o detección de glitches.
        """),
        ("Características Principales", """
• Rango de frecuencia: Desde DC hasta 300 MHz (en entradas A y B).

• Resolución de tiempo: Hasta 50 ps en modo de medición de intervalo de tiempo (T.I.) y 100 ps en mediciones estándar.

• Resolución de frecuencia: Hasta 12 dígitos/s; 1×10⁻¹¹ en 100 ms gracias al método de recuento recíproco interpolado.

• Zero-dead-time: Medición continua sin tiempo muerto, ideal para análisis de estabilidad tipo ADEV, TIE y TDEV, y para capturar micro-glitches de frecuencia.

• Mediciones back-to-back (BtB): Permite registrar series continuas de frecuencia o tiempo sin lag entre muestras, crucial para estudios de estabilidad.

• Entradas múltiples: Dos canales principales (A y B) con ajustes de impedancia, acoplamiento, filtrado analógico y digital, y un canal C opcional (según versión).

• Modulación Domain Analysis (TimeView™): Software opcional para análisis de comportamiento dinámico de frecuencia y detección de inestabilidades como jitter o glitch.

• Interfaz GPIB y USB: Hasta 4,000 resultados/s en GPIB y 10,000 mediciones/s en bloque; compatible con modo HP53132A para integración en sistemas existentes.

• Pantalla gráfica LCD retroiluminada: Permite visualizar resultados, histogramas, tendencia y alertas de límite en tiempo real.

• Trigger Hold-off programable: Evita falsas activaciones en presencia de rebotes o ruido.

• Capacidad de timestamping continuo: Ideal para cálculo de ADEV, TIE, TDEV y estudios térmicos o de arranque.
        """),
        ("¿Cómo realiza el CNT-91 las mediciones de frecuencia?", """
El CNT-91 mide frecuencia mediante conteo recíproco interpolado, una técnica que ofrece mucho mayor resolución que el conteo directo. En lugar de contar cuántos ciclos de la señal ocurren durante un intervalo fijo de tiempo (como hacen los contadores tradicionales), el CNT-91 mide el tiempo exacto entre eventos de la señal de entrada con interpolación analógica de los flancos. Luego, calcula la frecuencia como el inverso del período medido.

Por ejemplo, para señales de baja frecuencia, mide con precisión el período de la señal (el tiempo entre dos flancos) y lo invierte. Para señales más rápidas, cuenta múltiples ciclos y aplica interpolación de los flancos inicial y final, reduciendo el error.

La interpolación se realiza en hardware, cargando un capacitor con corriente constante desde el momento del flanco de la señal hasta el siguiente pulso del reloj de referencia. El voltaje resultante se convierte en una medida de subintervalo con una resolución de hasta 50 ps.

Este principio permite que la resolución de frecuencia mejore con el tiempo de medición, sin depender directamente de la frecuencia de entrada. A modo de ejemplo:

• En solo 100 ms, se alcanza una resolución relativa de 1×10⁻¹¹.
• Se pueden obtener 12 dígitos/s para frecuencias desde Hz hasta cientos de MHz.
        """),
        ("Medición de trenes de pulsos (Pulse Trains o Burst Mode)", """
Además de las funciones tradicionales de frecuencia y período, el CNT-91 incluye un modo especial para señales moduladas por pulsos (burst signals), común en osciladores que emiten trenes de pulsos o señales activadas periódicamente (como salidas sincronizadas en telecomunicaciones o GPS).

En este modo:

• El equipo sincroniza la medición con la aparición del burst, ignorando el tiempo de inactividad entre bursts.

• La frecuencia portadora dentro del burst se mide como si fuera una señal continua, sin necesidad de señal de armado externa.

• También se puede medir la frecuencia de repetición del burst (PRF - Pulse Repetition Frequency).

• El usuario define un retardo de sincronización (Sync Delay) y un número de ciclos a medir dentro del burst. Esto permite hacer mediciones precisas dentro del burst sin errores causados por transitorios o jitter de activación.

Este tipo de medición es vital para:

• Análisis de comportamiento de arranque (warm-up)
• Medición de estabilidad transitoria
• Validación de señales disparadas por eventos externos o controladores digitales
        """),
        ("Aplicaciones", """
• Mediciones de frecuencia de alta precisión: Ajuste, verificación y caracterización de osciladores.

• Calibración de equipos: Estándares de frecuencia, generadores, analizadores de espectro y sincronización de relojes.

• Investigación y desarrollo: Análisis de comportamiento de arranque, estabilidad a corto y largo plazo, análisis de PLLs y detección de glitches.

• Control de calidad: Verificación de parámetros de tiempo y frecuencia según especificaciones técnicas.

• Telecomunicaciones: Medición de parámetros de wander (TIE, TDEV) en relojes síncronos de red (ej. Stratum 1 a 3).

• Producción automatizada: Capacidad de medir conmutando rápidamente entre dispositivos (DUTs), detección de DUTs fallidos en 10 ms, y modo de presentación con límites de aceptación visuales.

• Medición de estabilidad a corto plazo (ADEV): Cálculo automático de σᵧ(τ) hasta 5000 s usando timestamping continuo.
        """)
    ]

def get_info_cnt91_resources():
    """
    Devuelve la lista de recursos oficiales (icono, texto, url) para la información del CNT-91.
    """
    return [
        ('\u2197', 'Página del Producto Pendulum', 'https://pendulum-instruments.com/products/frequency-counters-analyzers/cnt-91-91r/'),
        ('\U0001F4C4', 'Datasheet', 'https://pendulum-instruments.com/wp-content/uploads/2022/05/pendulum-cnt-91-91r_timer-counter-analyzer-calibrator.pdf'),
        ('\U0001F4D6', 'Manual de Usuario', 'https://pendulum-instruments.com/wp-content/uploads/2022/05/pendulum-cnt-91-91r_timer-counter-analyzer-calibrator.pdf'),
        ('\U0001F4BB', 'Manual de Programador', 'https://pendulum-instruments.com/wp-content/uploads/2022/05/CNT-90ph.pdf'),
    ] 