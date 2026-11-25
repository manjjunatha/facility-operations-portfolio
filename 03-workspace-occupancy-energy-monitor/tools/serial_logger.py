"""
Serial logger for the Workspace Occupancy and Energy Monitor.

Reads the ESP32's USB serial output and appends every DATA line to a CSV file
with the PC's date-time, ready to paste into the Raw_Data sheet of
Occupancy_Energy_Report.xlsx. ALERT lines are printed and saved to alerts.csv.

Usage:
    pip install pyserial
    python serial_logger.py --port COM5                    (Windows)
    python serial_logger.py --port /dev/ttyUSB0            (Linux)
    python serial_logger.py --port COM5 --out week1.csv

Stop with Ctrl+C. The file is flushed after every line, so nothing is lost.
"""
import argparse
import csv
import datetime as dt
import os
import sys

HEADER = ["Timestamp", "Room", "Occupied", "Light_On", "AC_On", "AC_Amps", "Waste_Flag"]


def parse_line(line):
    """Return ('data', row) / ('alert', row) / (None, None) for one serial line."""
    line = line.strip()
    if line.startswith("DATA,"):
        parts = line.split(",")
        if len(parts) != 8:
            return None, None
        _, _uptime, room, occ, light, ac, amps, waste = parts
        try:
            return "data", [room, int(occ), int(light), int(ac), float(amps), int(waste)]
        except ValueError:
            return None, None
    if line.startswith("ALERT,"):
        parts = line.split(",", 3)
        if len(parts) == 4:
            return "alert", [parts[2], parts[3]]
    return None, None


def open_csv(path, header):
    new = not os.path.exists(path) or os.path.getsize(path) == 0
    fh = open(path, "a", newline="", encoding="utf-8")
    w = csv.writer(fh)
    if new:
        w.writerow(header)
    return fh, w


def main():
    ap = argparse.ArgumentParser(description=__doc__, formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument("--port", required=True, help="Serial port, e.g. COM5 or /dev/ttyUSB0")
    ap.add_argument("--baud", type=int, default=115200)
    ap.add_argument("--out", default="occupancy_log.csv")
    ap.add_argument("--alerts", default="alerts.csv")
    args = ap.parse_args()

    try:
        import serial  # pyserial
    except ImportError:
        sys.exit("pyserial is not installed. Run:  pip install pyserial")

    data_fh, data_w = open_csv(args.out, HEADER)
    alert_fh, alert_w = open_csv(args.alerts, ["Timestamp", "Room", "Message"])
    print(f"Logging {args.port} @ {args.baud} -> {args.out}  (Ctrl+C to stop)")
    rows = 0
    try:
        with serial.Serial(args.port, args.baud, timeout=2) as ser:
            while True:
                raw = ser.readline()
                if not raw:
                    continue
                text = raw.decode("utf-8", errors="replace")
                kind, row = parse_line(text)
                ts = dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                if kind == "data":
                    data_w.writerow([ts] + row); data_fh.flush(); rows += 1
                    print(f"\r{rows} rows logged  (last: {ts} {row[0]})", end="")
                elif kind == "alert":
                    alert_w.writerow([ts] + row); alert_fh.flush()
                    print(f"\n*** {ts}  {row[0]}: {row[1]}")
                elif text.startswith("#"):
                    print("\n" + text.strip())
    except KeyboardInterrupt:
        print(f"\nStopped. {rows} rows saved to {args.out}")
    finally:
        data_fh.close(); alert_fh.close()


if __name__ == "__main__":
    main()
