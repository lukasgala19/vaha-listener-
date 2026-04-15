"""
Váha → kurzor (globálny listener)
===================================
Beží na pozadí. Keď váha pošle Print signál,
hodnota sa "napíše" tam kde je kurzor — Excel, Notepad, čokoľvek.

Inštalácia:
    pip install pyserial pyautogui

Spustenie:
    python vaha_listener.py

Ukončenie: Ctrl+C v terminály kde beží
"""

import serial
import serial.tools.list_ports
import pyautogui
import re
import sys
import time

# ─────────────────────────────────────────────
#  NASTAVENIA
# ─────────────────────────────────────────────
BAUD_RATE  = 9600
TIMEOUT    = 1

# Čo sa napíše za hodnotou — vyber jedno:
# ""        → iba číslo
# "\t"      → číslo + Tab  (posunie na ďalšiu bunku v Exceli)
# "\n"      → číslo + Enter
SUFFIX = "\t"

# Diagnostický mód: True = vypíše surové dáta, False = tichý beh
DIAGNOSTIC = False
# ─────────────────────────────────────────────


def zisti_port():
    porty = list(serial.tools.list_ports.comports())
    if not porty:
        print("⚠️  Žiadne COM porty. Je váha zapojená?")
        sys.exit(1)

    if len(porty) == 1:
        print(f"✅ Použijem port: {porty[0].device}  ({porty[0].description})")
        return porty[0].device

    print("\n📡 Dostupné COM porty:")
    for i, p in enumerate(porty):
        print(f"  [{i}] {p.device:10s}  {p.description}")
    idx = int(input("Vyber číslo: "))
    return porty[idx].device


def parsuj_hodnotu(riadok: str):
    """Extrahuje číslo z riadku váhy. Vracia string napr. '12.3456'"""
    riadok = riadok.strip()
    m = re.search(r"([+-]?\d[\d.,]+)", riadok)
    if m:
        return m.group(1).replace(",", ".")
    return None


def hlavna_slucka(port):
    print(f"\n🎯 Počúvam na {port} ({BAUD_RATE} baud)")
    print("   Stlač PRINT na váhach → hodnota sa objaví kde máš kurzor")
    print("   Ukonči: Ctrl+C\n")

    with serial.Serial(port, BAUD_RATE, timeout=TIMEOUT) as ser:
        buffer = ""
        while True:
            try:
                bajty = ser.read(128)
                if not bajty:
                    continue

                if DIAGNOSTIC:
                    raw = bajty.decode("ascii", errors="replace")
                    print(f"  RAW: {raw!r}")

                buffer += bajty.decode("ascii", errors="replace")

                # Spracuj kompletné riadky
                for sep in ["\r\n", "\n", "\r"]:
                    while sep in buffer:
                        riadok, buffer = buffer.split(sep, 1)
                        riadok = riadok.strip()
                        if not riadok:
                            continue

                        hodnota = parsuj_hodnotu(riadok)
                        if hodnota:
                            print(f"  ✅ Posielam: {hodnota}")
                            # Krátka pauza aby sa nestratilo focus
                            time.sleep(0.05)
                            pyautogui.typewrite(hodnota + SUFFIX, interval=0.02)
                        else:
                            print(f"  ⚠️  Nerozpoznaný formát: {riadok!r}")

            except KeyboardInterrupt:
                print("\nUkončené.")
                break


def main():
    print("=" * 50)
    print("  Váha → kurzor  (globálny listener)")
    print("=" * 50)

    # Diagnostický mód ako argument: python vaha_listener.py diag
    global DIAGNOSTIC
    if len(sys.argv) > 1 and sys.argv[1] == "diag":
        DIAGNOSTIC = True
        print("🔬 Diagnostický mód zapnutý\n")

    port = zisti_port()
    hlavna_slucka(port)


if __name__ == "__main__":
    main()
