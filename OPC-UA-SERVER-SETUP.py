import time
import datetime
import threading
from opcua import Server
from pymodbus.client.sync import ModbusSerialClient
import tkinter as tk
from tkinter import messagebox, scrolledtext
from PIL import Image, ImageTk
import os
import sys
import win32
# --- SMTP email ---
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import shutil
import math


# Ruta local y ruta en Google Drive
# RUTA CORRECTA PARA CSV (local al proyecto)
ruta_local = os.path.join(os.getcwd(), "registro_local.csv")

# RUTA DRIVE (sin cambios)
ruta_drive = r"G:\Mi unidad\Colab Notebooks\DATA\registro_datos.csv"

# Encabezados correctos
ENCABEZADOS = (
    "fecha,"
    "Presion_Cabeza,"
    "Presion_Separador,"
    "Presion_Choke,"
    "Presion_Desarenador,"
    "Presion_Aux1,"
    "Temp_Cabeza,"
    "Temp_Separador\n"
)


# <-- HORAS DE ENVÍO AUTOMÁTICO
horas_envio = ["09:00", "11:00", "15:30", "18:00", "05:00"]

# <-- CORREOS DESTINO
correos_destino = [
    "david.castillo@sipomega.com",
    "sensores1@sipomega.com",
    "jesus.dominguez@sipomega.com",
    "jesus.emm365@gmail.com"
]


# ---------------------------
# CONFIGURACIÓN SMTP GMAIL
# ---------------------------
SMTP_SERVER = "smtp.gmail.com"
SMTP_PORT = 587  # TLS
USER_EMAIL = "sipomega.inventario@gmail.com"   
USER_PASS = "doxg faok ivnl lvpf"     # <-- App Password de Gmail


def crear_csv_con_encabezado():
    try:
        # Si no existe, crear con encabezados
        if not os.path.exists(ruta_local):
            with open(ruta_local, "w", encoding="utf-8") as f:
                f.write(ENCABEZADOS)
            print("📄 CSV creado con encabezados (nuevo archivo).")
            return

        # Si existe pero está vacío, escribir encabezados
        if os.path.getsize(ruta_local) == 0:
            with open(ruta_local, "w", encoding="utf-8") as f:
                f.write(ENCABEZADOS)
            print("📄 Encabezados agregados (archivo vacío).")
            return

        # Si existe, verificar encabezado
        with open(ruta_local, "r", encoding="utf-8") as f:
            primera = f.readline().strip()

        if primera != ENCABEZADOS.strip():
            print("⚠ Encabezado incorrecto → Corrigiendo...")
            contenido = open(ruta_local, "r", encoding="utf-8").read()
            with open(ruta_local, "w", encoding="utf-8") as f:
                f.write(ENCABEZADOS + contenido)
        else:
            print("📄 Encabezado correcto, no se modifica.")

    except Exception as e:
        print(f"⚠ Error creando encabezado CSV: {e}")



def enviar_correo_con_csv():
    try:
        # Crear email multiparte
        msg = MIMEMultipart()
        msg["From"] = USER_EMAIL
        msg["To"] = ",".join(correos_destino)
        msg["Subject"] = "Reporte Automático SIPOMEGA"

        # Cuerpo del mensaje
        cuerpo = "Se adjunta archivo CSV con los datos del sistema SIPOMEGA."
        msg.attach(MIMEText(cuerpo, "plain"))

        # Adjuntar archivo CSV
        with open(ruta_local, "rb") as f:
            parte = MIMEBase("application", "octet-stream")
            parte.set_payload(f.read())
            encoders.encode_base64(parte)
            parte.add_header(
                "Content-Disposition",
                f"attachment; filename=registro_local.csv"
            )
            msg.attach(parte)

        # Envío del correo
        with smtplib.SMTP(SMTP_SERVER, SMTP_PORT) as server:
            server.ehlo()
            server.starttls()
            server.login(USER_EMAIL, USER_PASS)
            server.send_message(msg)

        print("📧 Correo enviado correctamente mediante Gmail.")

    except Exception as e:
        print(f"⚠ Error enviando correo mediante Gmail: {e}")

def programar_envio_correos():
    print("📬 Hilo de envío de correos activo.")
    while True:
        ahora = datetime.datetime.now().strftime("%H:%M")
        if ahora in horas_envio:
            enviar_correo_con_csv()
            time.sleep(60)
        time.sleep(5)




def respaldo_drive_cada_12h():
    while True:
        try:
            print("⏳ Esperando 60 minutos para respaldo...")
            time.sleep(600)  # 12 horas = 43,200 segundos

            shutil.copyfile(ruta_local, ruta_drive)
            print("✅ Respaldo copiado a Google Drive exitosamente.")

        except Exception as e:
            print(f"⚠️ Error al copiar al Drive: {e}")

# Variables globales
sistema_activo = False
lecturas_recientes = []

def iniciar_sistema():
    global sistema_activo

    try:
        bloquear_campos(True)  # Desactivar campos
        sistema_activo = True
        crear_csv_con_encabezado()

        threading.Thread(target=respaldo_drive_cada_12h, daemon=True).start()
        print("🕒 Hilo de respaldo automático iniciado.")
        threading.Thread(target=programar_envio_correos, daemon=True).start()

        estado_var.set("🟢 Sistema corriendo")

        # Obtener configuración desde GUI
        port = port_entry.get()
        baudrate = int(baudrate_entry.get())
        parity = parity_var.get()
        stopbits = int(stopbits_entry.get())
        bytesize = int(bytesize_entry.get())
        timeout = float(timeout_entry.get())
        unit_id = int(unitid_entry.get())
        endpoint = endpoint_entry.get()
        namespace_uri = namespace_entry.get()

        # --- Inicializa conexión Modbus ---
        print(f"\n🔌 Iniciando Modbus en {port} ({baudrate}bps, {parity}, {stopbits} stopbits)...")
        modbus = ModbusSerialClient(
            method='rtu',
            port=port,
            baudrate=baudrate,
            parity=parity,
            stopbits=stopbits,
            bytesize=bytesize,
            timeout=timeout
        )

        if not modbus.connect():
            print("❌ No se pudo conectar al puerto serial.")
            estado_var.set("❌ Error de conexión Modbus")
            bloquear_campos(False)
            return

        print("✅ Conexión Modbus establecida correctamente.")

        # --- Servidor OPC UA ---
        server = Server()
        server.set_endpoint(endpoint)
        idx = server.register_namespace(namespace_uri)
        sensores = server.nodes.objects.add_folder(idx, "Sensores")
        presion_c = sensores.add_variable(idx, "Presion_cabeza_psi", 0.0)
        presion_s = sensores.add_variable(idx, "Presion_separador_psi", 0.0)
        presion_ch = sensores.add_variable(idx, "Presion_choke_psi", 0.0)
        presion_d = sensores.add_variable(idx, "Presion_desarenador_psi", 0.0)
        presion_aux1 = sensores.add_variable(idx, "Presion_auxiliar1_psi", 0.0)

        presion_c_kg = sensores.add_variable(idx, "Presion_cabeza_kg/cm2", 0.0)
        presion_s_kg = sensores.add_variable(idx, "Presion_separador_kg/cm2", 0.0)
        presion_ch_kg = sensores.add_variable(idx, "Presion_choke_kg/cm2", 0.0)
        presion_d_kg = sensores.add_variable(idx, "Presion_desarenador_kg/cm2", 0.0)
        presion_aux1_kg = sensores.add_variable(idx, "Presion_auxiliar1_kg/cm2", 0.0)

        temperatura = sensores.add_variable(idx, "Temperatura_cabeza_Celsius", 0.0)
        temperatura2 = sensores.add_variable(idx, "Temperatura_separador_Celsius", 0.0)

        
        temperatura_f = sensores.add_variable(idx, "Temperatura_cabeza_Fahrenheit", 0.0)
        temperatura2_f = sensores.add_variable(idx, "Temperatura_separador_Fahrenheit", 0.0)
        

        presion_c.set_writable()
        presion_s.set_writable()
        presion_ch.set_writable()
        presion_d.set_writable()
        presion_aux1.set_writable()
        temperatura.set_writable()
        temperatura2.set_writable()
        temperatura_f.set_writable()
        temperatura2_f.set_writable()
        server.start()
        print(f"✅ Servidor OPC UA iniciado en {endpoint}")

        contador = 0

        while sistema_activo:
            contador += 1
            print(f"\n📡 Ciclo de lectura #{contador}")

            if not modbus.is_socket_open():
                print("⚠️ Reintentando conexión Modbus...")
                modbus.connect()

            rr = modbus.read_input_registers(address=0, count=9, unit=unit_id)
            if rr.isError():

                nan = math.nan

                presion_c.set_value(nan)
                presion_s.set_value(nan)
                presion_ch.set_value(nan)
                presion_d.set_value(nan)
                presion_aux1.set_value(nan)

                presion_c_kg.set_value(nan)
                presion_s_kg.set_value(nan)
                presion_ch_kg.set_value(nan)
                presion_d_kg.set_value(nan)
                presion_aux1_kg.set_value(nan)

                temperatura.set_value(nan)
                temperatura2.set_value(nan)
                temperatura_f.set_value(nan)
                temperatura2_f.set_value(nan)

                print(f"⚠️ Error al leer registros Modbus: {rr}")

            else:
                valor_presion0 = rr.registers[0]
                valor_presion1 = rr.registers[1]
                valor_presion2 = rr.registers[2]
                valor_presion3 = rr.registers[3]
                valor_presion4 = rr.registers[5]
                valor_temp1 = rr.registers[6]
                valor_temp2 = rr.registers[4]

                valor_presion_0 = ((valor_presion0 - 400) * (1500 - 0) / (2000 - 400)) + 0
                valor_presion_1 = ((valor_presion1 - 400) * (1500 - 0) / (2000 - 400)) + 0
                valor_presion_2 = ((valor_presion2 - 400) * (7000 - 0) / (2000 - 400)) + 0
                valor_presion_3 = ((valor_presion3 - 400) * (1500 - 0) / (2000 - 400)) + 0
                valor_presion_4 = ((valor_presion4 - 400) * (600 - 0) / (2000 - 400)) + 0

                valor_presion_5 = valor_presion_0 * 0.07
                valor_presion_6 = valor_presion_1 * 0.07
                valor_presion_7 = valor_presion_2 * 0.07
                valor_presion_8 = valor_presion_3 * 0.07
                valor_presion_9 = valor_presion_4 * 0.07


                valor_temp = ((valor_temp1 - 400) * (100 - 4) / (2000 - 400)) + 4
                valor_temp2 = ((valor_temp2 - 400) * (100 - 4) / (2000 - 400)) + 4

                valor_temp_f = (valor_temp*1.8)+32
                valor_temp2_f = (valor_temp2*1.8)+32
                #OPC-UPDATE
                presion_c.set_value(valor_presion_0)
                presion_s.set_value(valor_presion_1)
                presion_ch.set_value(valor_presion_2)
                presion_d.set_value(valor_presion_3)
                presion_aux1.set_value(valor_presion_4)

                presion_c_kg.set_value(valor_presion_5)
                presion_s_kg.set_value(valor_presion_6)
                presion_ch_kg.set_value(valor_presion_7)
                presion_d_kg.set_value(valor_presion_8)
                presion_aux1_kg.set_value(valor_presion_9)


                
                temperatura.set_value(valor_temp)
                temperatura2.set_value(valor_temp2)

                temperatura_f.set_value(valor_temp_f)
                temperatura2_f.set_value(valor_temp2_f)

                lectura_txt = (f"[{datetime.datetime.now():%H:%M:%S}] "
                               f" |  P_C={valor_presion_0:.2f} psi  | P_Ch={valor_presion_2:.2f} psi | P_D={valor_presion_3:.2f} psi | P_S={valor_presion_1:.2f} psi  |  P_aux1={valor_presion_4:.2f} psi\n "
				f" |  m_C={valor_presion0} mA  | m_Ch={valor_presion2} mA | m_D={valor_presion3} mA | m_S={valor_presion1} mA  |  m_aux1={valor_presion4} mA\n "
                               f"\t \t| T_C={valor_temp:.2f} °C | T_S={valor_temp2:.2f} °C" )
     
                print(lectura_txt)

                # Guardar lectura en lista y CSV

                #lecturas_recientes.append(lectura_txt)
                #if len(lecturas_recientes) > 50:
                #   lecturas_recientes.pop(0)

                #with open(ruta_local, "a") as f:
                #   f.write(f"{datetime.datetime.now()},{valor_presion_0},{valor_presion_1},{valor_presion_2},{valor_presion_3} ,{valor_presion_4}, {valor_temp}, {valor_temp2}\n")
                #ENCABEZADOS = "fecha,presion_0,presion_1,presion_2,presion_3,presion_4,temp_1,temp_2\n"

                # --- Tu lógica existente ---
                lecturas_recientes.append(lectura_txt)
                if len(lecturas_recientes) > 50:
                    lecturas_recientes.pop(0)

                with open(ruta_local, "a", encoding="utf-8") as f:
                    f.write(
                        f"{datetime.datetime.now():%d/%m/%Y %H:%M:%S},"
                        f"{valor_presion_0:.1f},"
                        f"{valor_presion_1:.1f},"
                        f"{valor_presion_2:.2f},"
                        f"{valor_presion_3:.2f},"
                        f"{valor_presion_4:.2f},"
                        f"{valor_temp:.2f},"
                        f"{valor_temp2:.2f}\n"
                    )

            time.sleep(5)

    except Exception as e:
        print(f"💥 Error general: {e}")
        estado_var.set("⚠️ Error en ejecución")
        bloquear_campos(False)

    finally:
        try:
            modbus.close()
            server.stop()
            print("✅ Sistema detenido correctamente.")
            estado_var.set("🔴 Sistema detenido")
        except:
            pass
        bloquear_campos(False)


def detener_sistema():
    global sistema_activo
    sistema_activo = False
    print("\n🛑 Deteniendo sistema...")

def bloquear_campos(bloquear=True):
    estado = "disabled" if bloquear else "normal"
    for e in campos_editables:
        e.config(state=estado)

def mostrar_lecturas():
    ventana = tk.Toplevel(root)
    ventana.title("📊 Lecturas en tiempo real")
    ventana.geometry("700x450")
    txt = scrolledtext.ScrolledText(ventana, wrap=tk.WORD, font=("Consolas", 10))
    txt.pack(fill="both", expand=True)
    txt.insert(tk.END, "\n".join(lecturas_recientes))
    txt.config(state="disabled")

# ---------- INTERFAZ ----------
root = tk.Tk()
root.title("SIPOMEGA | Modbus ↔ OPC UA")
root.geometry("610x400")
root.configure(bg="#f5f5f5")

# Marco principal dividido en 3 columnas
main_frame = tk.Frame(root, bg="#f5f5f5")
main_frame.pack(padx=15, pady=15, fill="both", expand=True)

# --- Columna 1: MODBUS ---
frame_modbus = tk.LabelFrame(main_frame, text="⚙️ Configuración Modbus", bg="#f5f5f5", padx=10, pady=10)
frame_modbus.grid(row=0, column=0, sticky="nsew", padx=10, pady=10)

def campo_modbus(label, default):
    tk.Label(frame_modbus, text=label, bg="#f5f5f5").pack()
    e = tk.Entry(frame_modbus)
    e.insert(0, default)
    e.pack(pady=2)
    return e

port_entry = campo_modbus("Puerto:", "COM11")
baudrate_entry = campo_modbus("Baudrate:", "9600")
parity_var = tk.StringVar(value="N")
tk.Label(frame_modbus, text="Paridad (N/E/O):", bg="#f5f5f5").pack()
tk.Entry(frame_modbus, textvariable=parity_var).pack(pady=2)
stopbits_entry = campo_modbus("Stop bits:", "1")
bytesize_entry = campo_modbus("Bytesize:", "8")
timeout_entry = campo_modbus("Timeout (s):", "3")
unitid_entry = campo_modbus("Unit ID:", "1")

# --- Columna 2: OPC UA ---
frame_opc = tk.LabelFrame(main_frame, text="🌐 Configuración OPC UA", bg="#f5f5f5", padx=10, pady=10)
frame_opc.grid(row=0, column=1, sticky="nsew", padx=10, pady=10)

def campo_opc(label, default):
    tk.Label(frame_opc, text=label, bg="#f5f5f5").pack()
    e = tk.Entry(frame_opc)
    e.insert(0, default)
    e.pack(pady=2)
    return e

endpoint_entry = campo_opc("Endpoint:", "opc.tcp://192.168.1.142:49320")
namespace_entry = campo_opc("Endpoint:", "http://sipomega.com/opcua/")

# --- Columna 3: CONTROLES ---
frame_botones = tk.LabelFrame(main_frame, text="🔘 Control", bg="#f5f5f5", padx=10, pady=10)
frame_botones.grid(row=0, column=2, sticky="nsew", padx=10, pady=10)

try:
    logo_img = Image.open("logo.png")
    logo_img = logo_img.resize((180, 80))
    logo_tk = ImageTk.PhotoImage(logo_img)
    tk.Label(frame_botones, image=logo_tk, bg="#f5f5f5").pack(pady=10)
except:
    tk.Label(frame_botones, text="SIPOMEGA", font=("Arial", 16, "bold"), bg="#f5f5f5").pack(pady=10)

estado_var = tk.StringVar(value="🔴 Sistema detenido")
tk.Label(frame_botones, textvariable=estado_var, font=("Arial", 11, "bold"), bg="#f5f5f5").pack(pady=10)

tk.Button(frame_botones, text="▶️ Iniciar Sistema",
          command=lambda: threading.Thread(target=iniciar_sistema, daemon=True).start(),
          bg="#4CAF50", fg="white", font=("Arial", 11, "bold"), width=20).pack(pady=5)

tk.Button(frame_botones, text="⏹️ Detener Sistema",
          command=detener_sistema,
          bg="#E53935", fg="white", font=("Arial", 11, "bold"), width=20).pack(pady=5)

tk.Button(frame_botones, text="📊 Ver Lecturas",
          command=mostrar_lecturas,
          bg="#1976D2", fg="white", font=("Arial", 11, "bold"), width=20).pack(pady=5)

# Guardar todos los campos editables para bloquearlos
campos_editables = [
    port_entry, baudrate_entry, stopbits_entry, bytesize_entry, timeout_entry,
    unitid_entry, endpoint_entry, namespace_entry
]

root.mainloop()
