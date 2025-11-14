from opcua import Client
import time
import datetime

endpoint = "opc.tcp://192.168.1.141:49320"   # IP de tu servidor OPC
client = Client(endpoint)

# Diccionario: nombre → nodeid
VARIABLES = {
    "PC_PSI":     "ns=2;i=2",
    "PE_PSI":   "ns=2;i=3",
    "PD_PSI":     "ns=2;i=4",
    "PS_PSI":   "ns=2;i=5",
    "PA1_PSI":     "ns=2;i=6",
    "PC_KGCM2":   "ns=2;i=7",
    "PE_KGCM2":    "ns=2;i=8",
    "PD_KGCM2":  "ns=2;i=9",
    "PS_KGCM2":    "ns=2;i=10",
    "PA1_KGCM2":  "ns=2;i=11",
    "T_C_C":      "ns=2;i=12",
    "T_S_C":      "ns=2;i=13",
    "T_C_F":      "ns=2;i=14",
    "T_S_F":       "ns=2;i=15"
}

try:
    print("Conectando...")
    client.connect()
    print("✔ Conectado\n")

    # Preparar nodos
    nodos = {nombre: client.get_node(nodeid) for nombre, nodeid in VARIABLES.items()}

    contador = 0
    while True:
        contador += 1
        timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        print(f"\n[{timestamp}] \n Ciclo de lectura #{contador}:")

        for nombre, nodo in nodos.items():
            try:
                valor = nodo.get_value()
                print(f"  {nombre}: {valor}")
            except Exception as e:
                print(f"  {nombre}: ERROR ({e})")

            with open("registro_datos.csv", "a") as f:
                    f.write(f"{datetime.datetime.now()},{nombre},{valor}\n")

        time.sleep(2)

except Exception as e:
    print("💥 Error general:", e)

finally:
    try:
        client.disconnect()
    except:
        pass
    print("❌ Cliente desconectado")
