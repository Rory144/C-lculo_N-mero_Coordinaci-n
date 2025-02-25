import warnings
import numpy as np
import pandas as pd
import matplotlib.pyplot as plt
from scipy.integrate import simps
from rich.console import Console
from rich.table import Table
from rich.panel import Panel
from rich.text import Text

# 🔹 Ocultar warnings innecesarios
warnings.filterwarnings("ignore", category=UserWarning)

# 🔹 Crear un objeto Console de rich
console = Console()

# 🔹 Banner de bienvenida
banner_text = Text("""
  _      _      _____ ____  ____    ____  _____   ____  ____  ____  ____  ____  _  _      ____  ____  _  _
 / \\  /|/ \\__/|/  __//  __\\/  _ \\  /  _ \\/  __/  /   _\\/  _ \\/  _ \\/  __\\/  _ \\/ \\/ \\  /|/  _ \/   _\/ \\/ \\  /|
 | |\\ ||| |\\/|||  \\  |  \\/|| / \\|  | | \\||  \\    |  /  | / \\|| / \\||  \\/|| | \\|| || |\\ ||| / \\||  /  | || |\\ ||
 | | \\||| |  |||  /_ |    /| \\_/|  | |_/||  /_   |  \\_ | \\_/|| \\_/||    /| |_/|| || | \\||| |-|||  \\_ | || | \\||
 \\_/  \\|\\_/  \\|\\____\\_/\\_\\____/  \\____/\\____\\  \\____/\\____/\\____/\\_/\\_\\____/\\_/\\_/  \\|\\_/ \\|\\____/\\_/\\_/  \\|
""", style="bold cyan")
console.print(banner_text)

console.print(Panel("🎉 [bold yellow]Bienvenido a la calculadora para el cálculo del número de coordinación![/bold yellow] 🎉\n"
                    "📂 Cargaremos los datos desde un archivo Excel y realizaremos los cálculos.",
                    title="⚙️ [bold cyan]Calculadora de Coordinación[/bold cyan]", style="bold green"))

def calcular_numero_coordinacion(archivo, hoja, r_min, r_max):
    """
    Calcula el número de coordinación a partir de datos de un archivo Excel.
    """
    # 🔹 Definir constantes
    densidad_metanol = 0.794e-24  # g/Å³ (densidad del metanol)
    numero_avogadro = 6.022e23    # moléculas/mol
    peso_molecular_metanol = 32   # g/mol
    densidad_numerica = (densidad_metanol * numero_avogadro) / peso_molecular_metanol

    # 🔹 Cargar los datos desde el archivo Excel
    df = pd.read_excel(archivo, sheet_name=hoja)
    
    # 🔹 Obtener columnas 'r' y 'g(r)'
    r = df.iloc[1:, 0].astype(float).values
    g_r = df.iloc[1:, 1].astype(float).values

    # 🔹 Calcular la función integrando g(r) * 4 * pi * r^2
    integrando = g_r * 4 * np.pi * r**2

    # 🔹 Filtrar los valores dentro del rango de integración
    mascara = (r >= r_min) & (r <= r_max)
    r_filtrado = r[mascara]
    integrando_filtrado = integrando[mascara]

    # 🔹 Mostrar tabla con valores filtrados
    table = Table(title="📊 Valores de r y g(r) * 4πr²", show_lines=True, header_style="bold blue")
    table.add_column("r", style="bold yellow", justify="center")
    table.add_column("g(r) * 4πr²", style="bold magenta", justify="center")
    
    for r_val, g_val in zip(r_filtrado, integrando_filtrado):
        table.add_row(f"{r_val:.2f}", f"{g_val:.2f}")

    console.print(table)

    # 🔹 Integrar usando la regla de Simpson
    integral = simps(integrando_filtrado, r_filtrado)
    
    # 🔹 Calcular número de coordinación
    numero_coordinacion = integral * densidad_numerica

    return numero_coordinacion

# 🔹 Parámetros del cálculo
archivo_excel = "Numero_coordinación.xlsx"
hoja_excel = "S0_CASSCF"
r_minimo = 0.0
r_maximo = 3.1

# 🔹 Calcular el número de coordinación
total_coordinacion = calcular_numero_coordinacion(archivo_excel, hoja_excel, r_minimo, r_maximo)

# 🔹 Mostrar el resultado con formato bonito
console.print(Panel(f"🔢 [bold cyan]Número de coordinación calculado:[/bold cyan] [bold yellow]{total_coordinacion:.4f}[/bold yellow]",
                    title="✅ Resultado", style="bold red"))

