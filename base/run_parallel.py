import subprocess
from concurrent.futures import ThreadPoolExecutor
from pathlib import Path
import time
import pandas as pd

# Configuración
START_LIC = 23036
END_LIC = 50000
NUM_THREADS = 2 # Número de hilos para mayor estabilidad
THREAD_DELAY = 10 #3600 Segundos entre inicio de hilos
BASE_DIR = Path(__file__).parent
OUTPUT_DIR = BASE_DIR / "output"


def initialize_excel_file(file_path):
    """Crea un archivo Excel válido con estructura inicial."""
    try:
        file_path.parent.mkdir(parents=True, exist_ok=True)
        if not file_path.exists():
            df = pd.DataFrame(columns=["Licencia", "Datos"])
            df.to_excel(file_path, index=False)
        return True
    except Exception as e:
        print(f"Error al inicializar {file_path}: {str(e)}")
        return False


def obtener_ultimo_procesado(hilo_num):
    progreso_file = BASE_DIR / f"progreso_hilo_{hilo_num}.txt"
    if progreso_file.exists():
        try:
            with open(progreso_file, "r") as file:
                return int(file.read().strip())
        except ValueError:
            print(f"⚠️ Error al leer el archivo de progreso del hilo {hilo_num}. Reiniciando desde el inicio.")
    return START_LIC - 1


def guardar_progreso(hilo_num, licencia_id):
    progreso_file = BASE_DIR / f"progreso_hilo_{hilo_num}.txt"
    with open(progreso_file, "w") as file:
        file.write(str(licencia_id))


def clean_robot_output(output_dir):
    """Limpia archivos XML de salida eliminando líneas con 'null'"""
    for file in output_dir.glob("output*.xml"):
        try:
            with open(file, 'r+', encoding='utf-8') as f:
                content = f.read()
                # Eliminar líneas con null
                cleaned = '\n'.join(line for line in content.split('\n') if 'null' not in line)
                f.seek(0)
                f.write(cleaned)
                f.truncate()
        except Exception as e:
            print(f"⚠️ Error limpiando {file}: {str(e)}")


def run_robot_process(range_start, range_end, output_file, hilo_num):
    ultimo_procesado = obtener_ultimo_procesado(hilo_num)
    if range_start <= ultimo_procesado < range_end:
        range_start = ultimo_procesado + 1

    if range_start > range_end:
        print(f"✅ Rango {range_start}-{range_end} ya completado, omitiendo... (Hilo {hilo_num})")
        return

    if not initialize_excel_file(output_file):
        print(f"❌ No se pudo inicializar {output_file} - omitiendo rango {range_start}-{range_end}")
        return

    command = [
        "robot",
        "--variable", f"START_LIC:{range_start}",
        "--variable", f"END_LIC:{range_end}",
        "--variable", f"EXCEL_FILE:{output_file}",
        "--outputdir", str(output_file.parent),
        "--loglevel", "DEBUG",
        "--exitonfailure",
        "task.robot"
    ]

    try:
        print(f"\n🚀 Iniciando proceso para rango {range_start}-{range_end} (Hilo {hilo_num})")
        result = subprocess.run(command, check=True, capture_output=True, text=True, timeout=86400)
        
        if "ERROR" in result.stdout or "FAIL" in result.stdout:
            print(f"⚠️ Advertencias/errores en {range_start}-{range_end} (Hilo {hilo_num}):")
            print(result.stdout[-500:])
            
        print(f"✅ Proceso completado para {range_start}-{range_end} (Hilo {hilo_num})")
        guardar_progreso(hilo_num, range_end)
        
    except subprocess.TimeoutExpired:
        print(f"⏰ Timeout para rango {range_start}-{range_end} (Hilo {hilo_num}).")
        guardar_progreso(hilo_num, range_start + (range_end - range_start) // 2)
        
    except subprocess.CalledProcessError as e:
        print(f"❌ Error en rango {range_start}-{range_end} (Hilo {hilo_num}):")
        print(f"STDOUT:\n{e.stdout[-500:] if e.stdout else 'Vacío'}")
        print(f"STDERR:\n{e.stderr[-500:] if e.stderr else 'Vacío'}")
        
        print(f"🔄 Reintentando rango {range_start}-{range_end}...")
        time.sleep(10)
        try:
            subprocess.run(command, check=True, capture_output=True, text=True, timeout=43200)
        except Exception as retry_e:
            print(f"❌ Error en reintento: {str(retry_e)}")
            
    except Exception as e:
        print(f"⚠️ Error inesperado en hilo {hilo_num}: {str(e)}")


def calculate_ranges():
    total = END_LIC - START_LIC + 1
    chunk_size = total // NUM_THREADS
    ranges = []
    for i in range(NUM_THREADS):
        start = START_LIC + i * chunk_size
        end = start + chunk_size - 1 if i < NUM_THREADS - 1 else END_LIC
        ranges.append((start, end))
    return ranges


if __name__ == "__main__":
    OUTPUT_DIR.mkdir(exist_ok=True)
    
    # Limpiar archivos de salida previos
    clean_robot_output(OUTPUT_DIR)
    
    license_ranges = calculate_ranges()

    with ThreadPoolExecutor(max_workers=NUM_THREADS) as executor:
        for i, (start, end) in enumerate(license_ranges):
            output_file = OUTPUT_DIR / f"resultados_hilo_{i + 1}.xlsx"
            executor.submit(run_robot_process, start, end, output_file, i + 1)
            if i < len(license_ranges) - 1:
                time.sleep(THREAD_DELAY)

    print("\n✅ Procesamiento completado. Verifica los resultados en la carpeta 'output'.")