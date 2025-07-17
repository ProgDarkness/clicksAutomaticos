import pyautogui
import time
import json
import os
import threading
import keyboard
import pytesseract
from PIL import Image
import sys

# Configuración de pytesseract (ajusta la ruta según tu instalación)
try:
    pytesseract.pytesseract.tesseract_cmd = r'C:\Program Files\Tesseract-OCR\tesseract.exe'  # Ruta común en Windows
    # Para Linux/macOS: normalmente se detecta automáticamente
except Exception as e:
    print(f"Error configurando Tesseract: {e}")

# Configuración básica
pyautogui.PAUSE = 0.5
pyautogui.FAILSAFE = True

CONFIG_FILE = "automatizacion_config.json"
ejecucion_automatica = False
intervalo = 30  # segundos entre ciclos completos

def verificar_dependencias():
    """Verifica e instala dependencias faltantes"""
    try:
        import pytesseract
    except ImportError:
        print("Instalando pytesseract...")
        os.system('pip install pytesseract')
        
    try:
        import keyboard
    except ImportError:
        print("Instalando keyboard...")
        os.system('pip install keyboard')
        
    try:
        from PIL import Image
    except ImportError:
        print("Instalando Pillow...")
        os.system('pip install pillow')

def mostrar_menu():
    print("\n--- MENÚ DE AUTOMATIZACIÓN ---")
    print("1. Agregar nueva acción")
    print("2. Ver acciones configuradas")
    print("3. Ejecutar automatización una vez")
    print("4. Iniciar ejecución automática (ciclo completo cada 30 seg)")
    print("5. Detener ejecución automática")
    print("6. Eliminar configuración")
    print("7. Salir")
    print("\nPresiona SUPR en cualquier momento para detener la ejecución")

def capturar_posicion(nombre_accion):
    print(f"\nPreparándose para capturar posición para: {nombre_accion}")
    print("Mueve el mouse a la posición deseada y presiona 's' para guardar (o 'q' para cancelar)...")
    
    while True:
        if keyboard.is_pressed('delete'):
            print("\nCaptura de posición cancelada por tecla SUPR")
            return None
        try:
            if keyboard.is_pressed('s'):
                x, y = pyautogui.position()
                print(f"Posición guardada: X={x}, Y={y} para '{nombre_accion}'")
                return {"x": x, "y": y, "nombre": nombre_accion}
            elif keyboard.is_pressed('q'):
                return None
        except:
            pass
        time.sleep(0.1)

def capturar_region(nombre_region):
    print(f"\nCapturando región para: {nombre_region}")
    print("Mueve el mouse a la esquina superior izquierda del área y presiona 's'")
    
    keyboard.wait('s')
    x1, y1 = pyautogui.position()
    print(f"Esquina superior izquierda: {x1}, {y1}")
    
    print("Mueve el mouse a la esquina inferior derecha del área y presiona 's'")
    keyboard.wait('s')
    x2, y2 = pyautogui.position()
    print(f"Esquina inferior derecha: {x2}, {y2}")
    
    return (x1, y1, x2 - x1, y2 - y1)

def agregar_accion():
    print("\n--- AGREGAR NUEVA ACCIÓN ---")
    print("Tipos de acción disponibles:")
    print("1. Clic")
    print("2. Escribir texto")
    print("3. Presionar tecla")
    print("4. Esperar texto en pantalla")
    
    tipo = input("Selecciona el tipo de acción (1-4): ")
    
    nombre = input("Nombre descriptivo para esta acción: ")
    
    if tipo == "1":
        posicion = capturar_posicion(nombre)
        if posicion:
            accion = {"tipo": "click", "posicion": posicion, "nombre": nombre}
            guardar_accion(accion)
    elif tipo == "2":
        texto = input("Texto a escribir: ")
        posicion = capturar_posicion(nombre)
        if posicion:
            accion = {"tipo": "escribir", "texto": texto, "posicion": posicion, "nombre": nombre}
            guardar_accion(accion)
    elif tipo == "3":
        tecla = input("Tecla a presionar (ej. 'enter', 'tab', 'esc'): ")
        accion = {"tipo": "tecla", "tecla": tecla, "nombre": nombre}
        guardar_accion(accion)
    elif tipo == "4":
        texto = input("Texto a esperar en pantalla: ")
        region = None
        if input("¿Quieres especificar una región específica? (s/n): ").lower() == 's':
            region = capturar_region(nombre)
        
        timeout = int(input("Tiempo máximo de espera en segundos (default 30): ") or "30")
        accion = {
            "tipo": "esperar_texto",
            "texto": texto,
            "region": region,
            "timeout": timeout,
            "nombre": nombre
        }
        guardar_accion(accion)
    else:
        print("Opción no válida")

def guardar_accion(nueva_accion):
    acciones = cargar_configuracion()
    acciones.append(nueva_accion)
    
    with open(CONFIG_FILE, 'w') as f:
        json.dump(acciones, f, indent=4)
    
    print(f"Acción '{nueva_accion['nombre']}' guardada correctamente")

def cargar_configuracion():
    if os.path.exists(CONFIG_FILE):
        with open(CONFIG_FILE, 'r') as f:
            return json.load(f)
    return []

def ver_acciones():
    acciones = cargar_configuracion()
    if not acciones:
        print("\nNo hay acciones configuradas")
        return
    
    print("\n--- ACCIONES CONFIGURADAS ---")
    for i, accion in enumerate(acciones, 1):
        print(f"{i}. {accion['nombre']} ({accion['tipo']})")
        if 'posicion' in accion:
            print(f"   Posición: X={accion['posicion']['x']}, Y={accion['posicion']['y']}")
        if 'texto' in accion:
            print(f"   Texto: '{accion['texto']}'")
        if 'tecla' in accion:
            print(f"   Tecla: '{accion['tecla']}'")
        if accion['tipo'] == "esperar_texto":
            print(f"   Texto a esperar: '{accion['texto']}'")
            if accion['region']:
                print(f"   Región: {accion['region']}")
            print(f"   Timeout: {accion['timeout']}s")

def esperar_texto(texto_esperado, timeout=30, region=None):
    """
    Espera hasta que el texto especificado aparezca en pantalla
    :param texto_esperado: Texto a buscar (case insensitive)
    :param timeout: Tiempo máximo de espera en segundos
    :param region: Tupla (x, y, width, height) para buscar en región específica
    :return: True si se encontró el texto, False si timeout
    """
    start_time = time.time()
    texto_esperado = texto_esperado.lower()
    
    while time.time() - start_time < timeout:
        if keyboard.is_pressed('delete'):
            print("\nEspera interrumpida por tecla SUPR")
            return False
            
        # Capturar pantalla o región
        try:
            screenshot = pyautogui.screenshot(region=region) if region else pyautogui.screenshot()
            
            # Mejorar contraste para OCR
            screenshot = screenshot.convert('L')  # Convertir a escala de grises
            # screenshot = screenshot.point(lambda x: 0 if x < 128 else 255)  # Binarización (opcional)
            
            # Usar OCR para extraer texto
            texto_en_pantalla = pytesseract.image_to_string(screenshot).lower()
            
            if texto_esperado in texto_en_pantalla:
                print(f"Texto encontrado: '{texto_esperado}'")
                return True
                
        except Exception as e:
            print(f"Error en OCR: {str(e)}")
            time.sleep(1)
            continue
        
        tiempo_restante = int(timeout - (time.time() - start_time))
        print(f"Esperando texto '{texto_esperado}'... ({tiempo_restante}s restantes)", end='\r')
        time.sleep(1)
    
    print("\nTiempo de espera agotado sin encontrar el texto")
    return False

def ejecutar_automatizacion():
    acciones = cargar_configuracion()
    if not acciones:
        print("\nNo hay acciones configuradas para ejecutar")
        return False
    
    print("\nPreparándose para ejecutar en 5 segundos...")
    print("Mueve el mouse a la esquina superior izquierda para cancelar")
    print("O presiona SUPR para cancelar inmediatamente")
    
    for i in range(5, 0, -1):
        if keyboard.is_pressed('delete'):
            print("\nEjecución cancelada por tecla SUPR")
            return False
        print(f"Tiempo restante: {i} segundos", end='\r')
        time.sleep(1)
    
    try:
        for i, accion in enumerate(acciones, 1):
            if keyboard.is_pressed('delete'):
                print("\nEjecución interrumpida por tecla SUPR")
                return False
                
            print(f"\nEjecutando acción {i}/{len(acciones)}: {accion['nombre']}")
            
            if accion['tipo'] == "click":
                pyautogui.click(accion['posicion']['x'], accion['posicion']['y'])
            elif accion['tipo'] == "escribir":
                pyautogui.click(accion['posicion']['x'], accion['posicion']['y'])
                pyautogui.write(accion['texto'])
            elif accion['tipo'] == "tecla":
                pyautogui.press(accion['tecla'])
            elif accion['tipo'] == "esperar_texto":
                region = tuple(accion['region']) if accion.get('region') else None
                if not esperar_texto(accion['texto'], accion['timeout'], region):
                    print(f"No se encontró el texto '{accion['texto']}' - deteniendo ejecución")
                    return False
            
            time.sleep(0.5)
        
        print("\nAutomatización completada")
        return True
    except pyautogui.FailSafeException:
        print("\nAutomatización cancelada por failsafe (mouse en esquina)")
        return False
    except Exception as e:
        print(f"\nError durante la ejecución: {str(e)}")
        return False

def ciclo_automatico():
    global ejecucion_automatica
    while ejecucion_automatica:
        if keyboard.is_pressed('delete'):
            print("\nTecla SUPR detectada - deteniendo ejecución automática")
            ejecucion_automatica = False
            break
            
        acciones = cargar_configuracion()
        if not acciones:
            print("\nNo hay acciones configuradas - deteniendo ejecución automática")
            ejecucion_automatica = False
            break
        
        print(f"\nIniciando ciclo de automatización - {time.strftime('%H:%M:%S')}")
        
        ejecucion_exitosa = ejecutar_automatizacion()
        
        if ejecucion_automatica and ejecucion_exitosa:
            print(f"\nCiclo completado. Esperando {intervalo} segundos para reiniciar...")
            
            for i in range(intervalo, 0, -1):
                if keyboard.is_pressed('delete') or not ejecucion_automatica:
                    ejecucion_automatica = False
                    print("\nEjecución automática detenida")
                    break
                print(f"Tiempo restante: {i} segundos", end='\r')
                time.sleep(1)
            
            print(" " * 30, end='\r')

def iniciar_ejecucion_automatica():
    global ejecucion_automatica
    if not ejecucion_automatica:
        acciones = cargar_configuracion()
        if not acciones:
            print("\nNo hay acciones configuradas para ejecutar")
            return
        
        print(f"\nIniciando ejecución automática con intervalo de {intervalo} segundos entre ciclos")
        print("Presiona SUPR en cualquier momento para detener")
        ejecucion_automatica = True
        
        hilo_automatico = threading.Thread(target=ciclo_automatico)
        hilo_automatico.daemon = True
        hilo_automatico.start()
    else:
        print("\nLa ejecución automática ya está en curso")

def detener_ejecucion_automatica():
    global ejecucion_automatica
    if ejecucion_automatica:
        ejecucion_automatica = False
        print("\nEjecución automática detenida")
    else:
        print("\nNo hay ejecución automática en curso")

def eliminar_configuracion():
    if os.path.exists(CONFIG_FILE):
        os.remove(CONFIG_FILE)
        print("Configuración eliminada")
    else:
        print("No hay configuración para eliminar")

def main():
    # Verificar e instalar dependencias faltantes
    verificar_dependencias()
    
    # Verificar Tesseract
    try:
        pytesseract.get_tesseract_version()
    except:
        print("\n¡ADVERTENCIA! Tesseract OCR no está instalado correctamente.")
        print("La función de detección de texto no funcionará sin Tesseract.")
        print("Descarga e instala Tesseract desde:")
        print("https://github.com/UB-Mannheim/tesseract/wiki")
        input("Presiona Enter para continuar sin detección de texto...")
    
    print("\nAUTOMATIZACIÓN INTELIGENTE CON DETECCIÓN DE TEXTO")
    print(f"Configuración: Ejecuta todas las acciones y espera {intervalo} segundos para repetir")
    print("** Presiona SUPR en cualquier momento para detener la ejecución automática **")
    
    try:
        while True:
            mostrar_menu()
            opcion = input("Selecciona una opción (1-7): ")
            
            if keyboard.is_pressed('delete'):
                detener_ejecucion_automatica()
                print("\nTecla SUPR detectada - volviendo al menú")
                continue
                
            if opcion == "1":
                agregar_accion()
            elif opcion == "2":
                ver_acciones()
            elif opcion == "3":
                ejecutar_automatizacion()
            elif opcion == "4":
                iniciar_ejecucion_automatica()
            elif opcion == "5":
                detener_ejecucion_automatica()
            elif opcion == "6":
                detener_ejecucion_automatica()
                eliminar_configuracion()
            elif opcion == "7":
                detener_ejecucion_automatica()
                print("Saliendo del programa...")
                break
            else:
                print("Opción no válida. Intenta de nuevo")
    except KeyboardInterrupt:
        detener_ejecucion_automatica()
        print("\nPrograma terminado por el usuario")

if __name__ == "__main__":
    main()