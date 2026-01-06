import serial
import time
import struct
import json
import os
from datetime import datetime, timedelta
from serial.tools import list_ports

# Prüfen auf Bibliotheken
try:
    from openpyxl import Workbook
except ImportError:
    print("\nFEHLER: Die Bibliothek 'openpyxl' fehlt.")
    print("Bitte installieren: pip install openpyxl\n")
    exit(1)

try:
    import matplotlib.pyplot as plt
    MATPLOTLIB_AVAILABLE = True
except ImportError:
    print("\nWARNUNG: 'matplotlib' fehlt. Keine grafische Ausgabe.")
    print("Bitte installieren: pip install matplotlib\n")
    MATPLOTLIB_AVAILABLE = False

# Name der Konfigurationsdatei
CONFIG_FILE = 'k204_config.json'

def parse_k204_packet(raw_data):
    """
    Analysiert das 45-Byte-Datenpaket des K204/HH309.
    """
    if len(raw_data) < 45:
        return None
    if raw_data[0] != 0x02 or raw_data[44] != 0x03:
        return None

    # --- Status Byte 1 ---
    status_byte_1 = raw_data[1]
    unit = '°C' if (status_byte_1 & 0x80) else '°F'
    
    # --- Status Bytes am Ende ---
    ol_chan_byte = raw_data[39]
    resolution_byte = raw_data[43]

    def get_bit(byte_val, bit_idx):
        return bool(byte_val & (1 << bit_idx))

    # --- Datenblöcke auslesen (Big Endian >hhhh) ---
    try:
        raw_current = struct.unpack('>hhhh', raw_data[7:15])
    except struct.error:
        return None

    # Container für die Ergebnisse
    temperatures = {}
    
    for i in range(4):
        channel_key = f"T{i+1}"
        
        # Auflösungs-Logik (Bit gesetzt = 0.1°C / Teiler 10)
        bit_set = get_bit(resolution_byte, i)
        divisor = 1.0 if bit_set else 10.0
        
        # Current Temp
        is_ol = get_bit(ol_chan_byte, i)
        val = raw_current[i] / divisor
        temperatures[channel_key] = 'OL' if is_ol else val

    decoded_data = {
        'unit': unit,
        'current_temperatures': temperatures
    }

    return decoded_data


def read_k204_data(port, baudrate=9600, timeout=3):
    """
    Liest und dekodiert das Datenpaket.
    """
    try:
        with serial.Serial(port, baudrate=baudrate, bytesize=serial.EIGHTBITS,
                           parity=serial.PARITY_NONE, stopbits=serial.STOPBITS_ONE,
                           timeout=timeout) as ser:
            ser.reset_input_buffer()
            ser.reset_output_buffer()
            time.sleep(0.1) 

            ser.write(b'\x41')
            time.sleep(0.5)

            raw_data = ser.read(45)
            return parse_k204_packet(raw_data)

    except Exception as e:
        print(f"Fehler beim Lesen: {e}")
        return None

# --- Konfigurations-Funktionen ---

def load_config():
    """Lädt Konfiguration (Kanäle, Settings) oder erstellt Standards."""
    defaults = {
        "channels": {"T1": "Kanal 1", "T2": "Kanal 2", "T3": "Kanal 3", "T4": "Kanal 4"},
        "settings": {"cycles": 0, "prefix": "messung", "interval": 1}
    }
    
    if os.path.exists(CONFIG_FILE):
        try:
            with open(CONFIG_FILE, 'r', encoding='utf-8') as f:
                data = json.load(f)
                if "channels" not in data:
                    new_conf = defaults.copy()
                    for k in ["T1", "T2", "T3", "T4"]:
                        if k in data:
                            new_conf["channels"][k] = data[k]
                    return new_conf
                
                if "settings" in data:
                    defaults["settings"].update(data["settings"])
                if "channels" in data:
                    defaults["channels"].update(data["channels"])
                return defaults
        except Exception as e:
            print(f"Fehler beim Laden der Konfiguration: {e}")
            return defaults
    return defaults

def save_config(config):
    """Speichert die Konfiguration in einer JSON-Datei."""
    try:
        with open(CONFIG_FILE, 'w', encoding='utf-8') as f:
            json.dump(config, f, indent=4)
        print("-> Konfiguration gespeichert.")
    except Exception as e:
        print(f"Fehler beim Speichern der Konfiguration: {e}")

def setup_menu():
    """
    Menü für Port, Config (Prefix, Zyklen, Kanäle) und Intervall.
    """
    print("--- VOLTCRAFT K204 Messgerät-Setup (XLSX & Plot) ---")
    
    config = load_config()
    
    # 1. COM-Port auswählen
    print("\nVerfügbare serielle Ports werden gesucht...")
    ports = list_ports.comports()
    if not ports:
        print("Keine seriellen Ports gefunden. Bitte Gerät anschließen.")
        return None, None, None, None, None
    
    ports_list = [port.device for port in ports]
    for i, port in enumerate(ports_list):
        print(f"  {i+1}: {port}")
    
    selected_port = None
    while selected_port is None:
        try:
            choice_input = input(f"Bitte Port-Nummer auswählen (1-{len(ports_list)}): ")
            choice = int(choice_input)
            if 1 <= choice <= len(ports_list):
                selected_port = ports_list[choice-1]
        except ValueError:
            pass
            
    # 2. Einstellungen bearbeiten
    print("\n--- Einstellungen ---")
    current_prefix = config["settings"].get("prefix", "messung")
    current_cycles = config["settings"].get("cycles", 0)
    current_interval = config["settings"].get("interval", 1)
    
    print(f"Aktueller Datei-Prefix: {current_prefix}")
    print(f"Aktuelle Zyklen (0=Endlos): {current_cycles}")
    
    edit = input("Einstellungen ändern? (j/n) [n]: ").lower()
    if edit == 'j':
        new_prefix = input(f"Neuer Datei-Prefix [{current_prefix}]: ").strip()
        if new_prefix:
            config["settings"]["prefix"] = new_prefix
            
        cycles_input = input(f"Anzahl Messzyklen (0=Endlos) [{current_cycles}]: ").strip()
        if cycles_input:
            try:
                config["settings"]["cycles"] = int(cycles_input)
            except ValueError:
                print("Ungültige Zahl.")

        interval_input = input("\nMessintervall in Sekunden (z.B. 1.5): ")
        interval = float(interval_input)
        if interval <= 0: interval = None
        if interval_input:
            try:
                config["settings"]["interval"] = int(interval_input)
            except ValueError:
                print("Ungültige Zahl.")

        print("\nKanal-Beschreibungen:")
        for key in ["T1", "T2", "T3", "T4"]:
            curr = config["channels"].get(key, "")
            new_name = input(f"  {key} [{curr}]: ").strip()
            if new_name:
                config["channels"][key] = new_name
        
        save_config(config)

    cycles = config["settings"]["cycles"]
    prefix = config["settings"]["prefix"]
    interval = config["settings"]["interval"]
    
    timestamp_str = datetime.now().strftime("%Y%m%d_%H%M%S")
    xlsx_filename = f"{prefix}_{timestamp_str}.xlsx"

#    interval = None
#    while interval is None:
#        try:
#            interval_input = input("\nMessintervall in Sekunden (z.B. 1.5): ")
#            interval = float(interval_input)
#            if interval <= 0: interval = None
#        except ValueError:
#            pass
            
    return selected_port, cycles, interval, config, xlsx_filename


if __name__ == '__main__':
    
    port, num_cycles, interval_sec, config, filename = setup_menu()
    
    if port is None:
        print("Setup abgebrochen.")
    else:
        print(f"\nStarte Messung auf {port}...")
        print(f"Speichere Excel-Datei: {filename}")
        print(f"Zyklen: {'Endlos' if num_cycles == 0 else num_cycles}")
        print(f"Intervall: {interval_sec}s")

        # --- EXCEL SETUP ---
        wb = Workbook()
        ws = wb.active
        ws.title = "Messdaten"

        header = ["Datum Zeit", "Laufzeit", "Sekunden"]
        channels = config["channels"]
        channel_keys = ["T1", "T2", "T3", "T4"]
        for key in channel_keys:
            header.append(f"{key} ({channels.get(key, '')})")
        header.append("Einheit")
        ws.append(header)
        
        ws.column_dimensions['A'].width = 20
        ws.column_dimensions['B'].width = 12
        for col in ['D', 'E', 'F', 'G']:
            ws.column_dimensions[col].width = 15

        try:
            wb.save(filename)
        except PermissionError:
            print(f"FEHLER: Kein Schreibzugriff auf {filename}. Ist die Datei offen?")
            exit(1)

        # --- PLOT SETUP ---
        if MATPLOTLIB_AVAILABLE:
            plt.ion() # Interactive Mode on
            fig, ax = plt.subplots(figsize=(10, 6))
            ax.set_title(f"VOLTCRAFT K204 - Live Daten ({filename})")
            ax.set_xlabel("Laufzeit [s]")
            ax.set_ylabel("Temperatur")
            ax.grid(True)
            
            # Linien initialisieren
            lines = []
            plot_colors = ['red', 'blue', 'green', 'orange']
            for i, key in enumerate(channel_keys):
                label_text = f"{key}: {channels.get(key, '')}"
                line, = ax.plot([], [], label=label_text, color=plot_colors[i], linewidth=1.5)
                lines.append(line)
            
            ax.legend(loc='upper left')
            
            # Listen für Plot-Daten
            x_data = []
            y_data = {k: [] for k in channel_keys}

        start_time = datetime.now()
        cycle_count = 0
        
        try:
            while True:
                if num_cycles > 0 and cycle_count >= num_cycles:
                    break
                    
                cycle_count += 1
                loop_start = time.time()
                
                current_time = datetime.now()
                elapsed = current_time - start_time
                elapsed_str = str(elapsed).split('.')[0]
                total_seconds = elapsed.total_seconds()
                
                full_data = read_k204_data(port)
                
                if full_data:
                    temps = full_data['current_temperatures']
                    unit = full_data['unit']
                    
                    # --- EXCEL ---
                    row = [
                        current_time.strftime("%Y-%m-%d %H:%M:%S"),
                        elapsed_str,
                        round(total_seconds, 1),
                        temps['T1'],
                        temps['T2'],
                        temps['T3'],
                        temps['T4'],
                        unit
                    ]
                    ws.append(row)
                    wb.save(filename)
                    
                    # --- PLOT UPDATE ---
                    if MATPLOTLIB_AVAILABLE:
                        x_data.append(total_seconds)
                        
                        # Einheit im Plot aktualisieren (nur einmal nötig eigentlich)
                        ax.set_ylabel(f"Temperatur [{unit}]")

                        for i, key in enumerate(channel_keys):
                            val = temps[key]
                            # OL ignorieren für Plot (None wird von matplotlib nicht geplottet -> Lücke)
                            if val == 'OL':
                                y_data[key].append(None)
                            else:
                                y_data[key].append(val)
                            
                            lines[i].set_data(x_data, y_data[key])
                        
                        # Achsen skalieren
                        ax.relim()
                        ax.autoscale_view()
                        
                        # UI Update - kurzes Pause-Event reicht
                        plt.pause(0.001)

                    # --- KONSOLE ---
                    def fmt(val):
                        if val == 'OL': return 'OL'
                        return f"{val:.1f}"

                    print(f"#{cycle_count:<4} | {elapsed_str} | "
                          f"T1: {fmt(temps['T1'])} | T2: {fmt(temps['T2'])} | "
                          f"T3: {fmt(temps['T3'])} | T4: {fmt(temps['T4'])}")
                else:
                    print(f"#{cycle_count:<4} | Keine Daten empfangen.")

                # Timing
                process_duration = time.time() - loop_start
                sleep_time = interval_sec - process_duration
                
                if num_cycles == 0 or cycle_count < num_cycles:
                    if sleep_time > 0:
                        time.sleep(sleep_time)

        except KeyboardInterrupt:
            print("\nMessung durch Benutzer (Strg+C) beendet.")
        except PermissionError:
             print("\nFEHLER beim Speichern: Datei ist möglicherweise in Excel geöffnet!")
        finally:
            if MATPLOTLIB_AVAILABLE:
                plt.ioff() # Interactive Mode off
                print("Fenster schließen, um Programm vollständig zu beenden...")
                plt.show() # Fenster offen lassen am Ende
            
            try:
                wb.save(filename)
                print("Datei erfolgreich gespeichert.")
            except:
                pass
        
        print("Programm beendet.")