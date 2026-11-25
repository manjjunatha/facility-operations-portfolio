# Workspace Occupancy and Energy Monitor
**ESP32 · PIR / LDR / CT-clamp sensors · Python logger · MS Excel**

An ESP32 watches up to three rooms. A PIR sensor detects whether each room is occupied, an LDR detects whether its lights are on, and a non-invasive CT clamp measures whether its AC is drawing current. If a room stays **empty with lights or AC on for more than 10 minutes**, the ESP32 sounds a buzzer, lights an LED, and logs an alert. Every minute it logs each room's state over USB; a Python script saves this to CSV, and an Excel report turns it into utilisation %, hours of waste, kWh and ₹ lost per room.

## Folder layout
```
firmware/OccupancyEnergyMonitor/OccupancyEnergyMonitor.ino   ESP32 sketch (Arduino IDE)
tools/serial_logger.py                                        saves serial output to CSV
excel/Occupancy_Energy_Report.xlsx                            dashboard (paste CSV into Raw_Data)
sample_data/occupancy_log_sample.csv                          simulated 5-day log used in the report
```

## Bill of materials (per 3 rooms)
| Qty | Part | Approx. price (Rs) |
|---|---|---|
| 1 | ESP32 DevKit V1 (38-pin) | 400 |
| 3 | HC-SR501 PIR motion sensor | 3 × 70 |
| 3 | LDR (5 mm) + 10 kΩ resistor | 3 × 10 |
| 3 | SCT-013-030 CT clamp (30 A / 1 V, built-in burden) | 3 × 450 |
| 3 | 3.5 mm audio jack breakout (for CT plug) | 3 × 30 |
| 6 | 10 kΩ resistors (CT bias divider) + 3 × 10 µF capacitors | 30 |
| 1 | Active buzzer 5 V | 20 |
| 1 | Breadboard / perfboard, jumper wires, USB cable | 150 |

*Prices are rough Robu.in / local-market estimates; check current listings.*

## Wiring
| Room | PIR OUT | LDR (analog) | CT clamp (analog) |
|---|---|---|---|
| Room 101 - Classroom | GPIO 27 | GPIO 32 | GPIO 33 |
| Lab 2 - Electronics | GPIO 26 | GPIO 34 | GPIO 35 |
| Faculty Room | GPIO 25 | GPIO 36 (VP) | GPIO 39 (VN) |

Buzzer + → GPIO 14 (– to GND). Status LED = on-board GPIO 2.

- **PIR (HC-SR501):** VCC → 5 V (VIN), GND → GND, OUT → GPIO. Output is 3.3 V, safe for the ESP32. Set the module's time-delay pot to minimum; the firmware handles the 5-minute hold.
- **LDR:** 3.3 V → LDR → GPIO → 10 kΩ → GND. More light = higher voltage. Mount it facing the room's tube lights, away from windows.
- **CT clamp (SCT-013-030):** clip it around **one** wire (live only) of the AC's supply. *It is non-invasive, so no mains wiring is opened.* Make a 1.65 V bias point with two 10 kΩ resistors in series between 3.3 V and GND, plus 10 µF from the midpoint to GND. Connect one lead of the clamp's 3.5 mm plug to this bias point and the other lead to the room's CT GPIO. The GPIO then sees 1.65 V ± the AC signal.

> ⚠️ Only clip the CT clamp around an insulated wire, for example inside the AC's isolator box with an electrician's help. Never open live mains wiring yourself.

Only ADC1 pins (32–39) are used, so Wi-Fi can be switched on without breaking analog reads.

## Firmware
1. Arduino IDE → Boards Manager → install **esp32 by Espressif**. Select **ESP32 Dev Module**.
2. Open `firmware/OccupancyEnergyMonitor/OccupancyEnergyMonitor.ino` and upload.
3. Open the Serial Monitor at **115200** baud. After a 30 s PIR warm-up you will see lines like:
```
DATA,90,Room 101 - Classroom,1,1,1,6.12,0
ALERT,1512,Room 101 - Classroom,Lights + AC left ON in empty room
```
`DATA,uptime_s,room,occupied,light_on,ac_on,ac_amps,waste`

**Settings at the top of the sketch:**
| Constant | Default | Meaning |
|---|---|---|
| `OCCUPANCY_TIMEOUT_MS` | 5 min | no motion this long ⇒ room empty |
| `WASTE_GRACE_MS` | 10 min | empty + ON this long ⇒ waste alert |
| `LOG_INTERVAL_MS` | 60 s | logging interval (set the same in Excel Settings) |
| `LIGHT_ON_MV` | 1500 mV | LDR threshold, calibrate on site |
| `AC_ON_AMPS` | 0.8 A | CT threshold, calibrate on site |
| `ENABLE_WIFI` | 0 | 1 = live status web page at `http://<ip>/` and `/csv` |

### Calibration (10 minutes)
1. Temporarily add `Serial.println(analogReadMilliVolts(32));` in `loop()`. Note the mV reading with lights OFF and ON. Set `LIGHT_ON_MV` midway between them.
2. With the AC off, then running, read `ac_amps` in the DATA lines. Idle/standby is usually under 0.3 A and a running 1.5-ton split AC 5–7 A. Keep `AC_ON_AMPS` at about 0.8 A or set it midway.

## Logging to CSV
```
pip install pyserial
python tools/serial_logger.py --port COM5 --out occupancy_log.csv
```
Each DATA line is saved with the PC's date-time; alerts go to `alerts.csv`. Leave it running for a week.

## Excel report
1. Open `excel/Occupancy_Energy_Report.xlsx`, delete the sample rows in **Raw_Data**, and paste your CSV rows from A5.
2. In **Settings**, set *Log interval* = 1 (minute) and check the room names, lighting loads (W), tariff (₹/kWh) and working hours.
3. **Dashboard** shows, per room: occupied hours, working-hours utilisation %, hours lights/AC ran in an empty room, kWh and ₹ wasted, 30-day projection, and an *Under-used* / *Energy waste* status. Hour-by-hour heatmaps show when rooms are used and when energy is wasted, and a daily chart shows the trend.

**Energy formulas:** lights kWh = empty-and-lights-on hours × lighting load ÷ 1000; AC kWh = Σ measured amps while empty × 230 V × 0.9 PF × interval ÷ 1000.

### Sample results (simulated week, 16–20 Nov 2026)
The supplied sample data is simulated with the same rules as the firmware. It shows 47.7 kWh wasted in 5 days (≈ ₹380, ≈ ₹2,290 projected per 30 days). Most of the waste came from the Lab's AC being left on overnight on Wednesday. The Lab was used only 44% of working hours, so it is flagged *Under-used*.

## Testing done
- Firmware compiled on a PC against stub Arduino headers and run in a simulated timeline. A room left with lights and AC on after the last motion went *empty* 5 minutes later, raised exactly one alert after the 10-minute grace period, and reported 2.12 A for a 100 mV-peak CT signal, as calculated.
- Logger parser tested on that firmware output (117 data lines, 1 alert, bad lines ignored).
- Every Excel figure was cross-checked against an independent pandas calculation on the same CSV. All matched.
