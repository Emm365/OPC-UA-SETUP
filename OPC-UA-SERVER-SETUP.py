import time
import datetime
import threading
from opcua import Server
from pymodbus.client.sync import ModbusSerialClient
import tkinter as tk
from tkinter import messagebox, scrolledtext
from PIL import Image, ImageTk

# Variables globales
sistema_activo = False
lecturas_recientes = []

def iniciar_sistema():
    global sistema_activo

    try:
        bloquear_campos(True)  # Desactivar campos
        sistema_activo = True
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
        sensores = server.nodes.objects.add_object(idx, "Sensores")
        presion_c = sensores.add_variable(idx, "Presion_cabeza", 0.0)
        presion_s = sensores.add_variable(idx, "Presion_separador", 0.0)
        presion_ch = sensores.add_variable(idx, "Presion_choke", 0.0)
        presion_d = sensores.add_variable(idx, "Presion_desarenador", 0.0)
        presion_aux1 = sensores.add_variable(idx, "Presion_auxiliar_1", 0.0)
        temperatura = sensores.add_variable(idx, "Temperatura_cabeza", 0.0)
        temperatura2 = sensores.add_variable(idx, "Temperatura_separador", 0.0)
        presion_c.set_writable()
        presion_s.set_writable()
        presion_ch.set_writable()
        presion_d.set_writable()
        presion_aux1.set_writable()
        temperatura.set_writable()
        temperatura2.set_writable()
        server.start()
        print(f"✅ Servidor OPC UA iniciado en {endpoint}")

        contador = 0

        while sistema_activo:
            contador += 1
            print(f"\n📡 Ciclo de lectura #{contador}")

            if not modbus.is_socket_open():
                print("⚠️ Reintentando conexión Modbus...")
                modbus.connect()

            rr = modbus.read_input_registers(address=0, count=8, unit=unit_id)
            if rr.isError():
                print(f"⚠️ Error al leer registros Modbus: {rr}")
            else:
                valor_presion0 = rr.registers[0]
                valor_presion1 = rr.registers[1]
                valor_presion2 = rr.registers[2]
                valor_presion3 = rr.registers[3]
                valor_presion4 = rr.registers[5]
                valor_temp1 = rr.registers[7]
                valor_temp2 = rr.registers[4]

                valor_presion_0 = ((valor_presion0 - 400) * (1500 - 0) / (2000 - 400)) + 0
                valor_presion_1 = ((valor_presion1 - 400) * (600 - 0) / (2000 - 400)) + 0
                valor_presion_2 = ((valor_presion2 - 400) * (600 - 0) / (2000 - 400)) + 0
                valor_presion_3 = ((valor_presion3 - 400) * (1500 - 0) / (2000 - 400)) + 0
                valor_presion_4 = ((valor_presion4 - 400) * (1500 - 0) / (2000 - 400)) + 0
                valor_temp = ((valor_temp1 - 400) * (100 - 4) / (2000 - 400)) + 4
                valor_temp_2 = ((valor_temp2 - 400) * (100 - 4) / (2000 - 400)) + 4

                presion_c.set_value(valor_presion_0)
                presion_s.set_value(valor_presion_1)
                presion_ch.set_value(valor_presion_2)
                presion_d.set_value(valor_presion_3)
                presion_aux1.set_value(valor_presion4)
                
                temperatura.set_value(valor_temp)
                temperatura2.set_value(valor_temp_2)

                lectura_txt = (f"[{datetime.datetime.now():%H:%M:%S}] "
                               f"P_C={valor_presion_0:.2f} psi | P_Ch={valor_presion_2:.2f} psi | P_D={valor_presion_3:.2f} psi | P_S={valor_presion_1:.2f} psi | P_aux1={valor_presion_4:.2f} psi \n"
                               f" \t T_C={valor_temp:.2f} °C | T_S={valor_temp_2:.2f} °C | "  )
                print(lectura_txt)

                # Guardar lectura en lista y CSV
                lecturas_recientes.append(lectura_txt)
                if len(lecturas_recientes) > 50:
                    lecturas_recientes.pop(0)

                with open("registro_datos.csv", "a") as f:
                    f.write(f"{datetime.datetime.now()},{valor_presion_0},{valor_presion_1},{valor_presion_2},{valor_presion_3}, {valor_temp}, {valor_temp2}\n")

            time.sleep(2)

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

port_entry = campo_modbus("Puerto:", "COM15")
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

endpoint_entry = campo_opc("Endpoint:", "opc.tcp://192.168.0.183:49320")
namespace_entry = campo_opc ("Namespace URI","http://sipomega.com/opcua/")

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
