/*
 * Workspace Occupancy and Energy Monitor  (ESP32, Arduino framework)
 * ---------------------------------------------------------------------------
 * Per room:  PIR motion sensor (HC-SR501)     -> is the room occupied?
 *            LDR light sensor (voltage divider) -> are the lights on?
 *            SCT-013-030 CT clamp (non-invasive) -> is the AC drawing current?
 *
 * The ESP32 flags "energy waste" when a room has been EMPTY but its lights
 * or AC have stayed ON for longer than WASTE_GRACE_MS. It sounds a buzzer,
 * prints an ALERT line, and logs one CSV line per room every LOG_INTERVAL_MS
 * over USB serial. tools/serial_logger.py saves those lines to a CSV file
 * that is pasted into Occupancy_Energy_Report.xlsx.
 *
 * Optional (ENABLE_WIFI 1): joins Wi-Fi and serves a live status page at
 * http://<esp32-ip>/  and the latest readings as CSV at /csv.
 *
 * Board: "ESP32 Dev Module" (Arduino-ESP32 core 2.x or 3.x)
 * Only ADC1 pins (32-39) are used for analog inputs so Wi-Fi can stay on.
 */

#include <Arduino.h>

#define ENABLE_WIFI 0           // 1 = enable live web page (set SSID/PASS below)

#if ENABLE_WIFI
#include <WiFi.h>
#include <WebServer.h>
const char *WIFI_SSID = "YOUR_WIFI_NAME";
const char *WIFI_PASS = "YOUR_WIFI_PASSWORD";
WebServer server(80);
#endif

// ----------------------------- configuration -------------------------------
const uint8_t NUM_ROOMS = 3;

struct RoomPins {
  const char *name;
  uint8_t pirPin;   // digital input
  uint8_t ldrPin;   // ADC1 analog input
  uint8_t ctPin;    // ADC1 analog input (CT clamp, biased at ~1.65 V)
};

const RoomPins ROOMS[NUM_ROOMS] = {
  {"Room 101 - Classroom", 27, 32, 33},
  {"Lab 2 - Electronics",  26, 34, 35},
  {"Faculty Room",         25, 36, 39},
};

const uint8_t BUZZER_PIN = 14;
const uint8_t STATUS_LED = 2;              // on-board LED

const uint32_t OCCUPANCY_TIMEOUT_MS = 5UL * 60UL * 1000UL;   // no motion for 5 min => empty
const uint32_t WASTE_GRACE_MS       = 10UL * 60UL * 1000UL;  // empty + ON for 10 min => waste
const uint32_t LOG_INTERVAL_MS      = 60UL * 1000UL;         // one CSV line per room per minute
const uint32_t SENSE_INTERVAL_MS    = 1000UL;                // sensor poll rate

// Calibrate these two on site (see README "Calibration"):
const uint16_t LIGHT_ON_MV  = 1500;   // LDR divider output above this => lights ON
const float    AC_ON_AMPS   = 0.8f;   // CT RMS current above this => AC ON

// SCT-013-030 has a built-in burden resistor: 1 V RMS at 30 A RMS
const float CT_AMPS_PER_VOLT = 30.0f;
const uint16_t CT_SAMPLES    = 400;   // ~4 mains cycles at 50 Hz with 100 us spacing

// ------------------------------- state -------------------------------------
struct RoomState {
  uint32_t lastMotionMs = 0;
  bool     everMotion   = false;
  bool     occupied     = false;
  bool     lightOn      = false;
  bool     acOn         = false;
  float    acAmps       = 0.0f;
  uint32_t wasteStartMs = 0;      // 0 = not currently empty-and-on
  bool     waste        = false;
  uint32_t wasteAlerts  = 0;
};

RoomState state[NUM_ROOMS];
uint32_t lastSenseMs = 0;
uint32_t lastLogMs   = 0;

// ------------------------------- sensors -----------------------------------
float readCtAmps(uint8_t pin) {
  // Two passes: find the DC bias (mean), then the RMS of the AC component.
  static uint16_t buf[CT_SAMPLES];
  uint32_t sum = 0;
  for (uint16_t i = 0; i < CT_SAMPLES; i++) {
    buf[i] = analogReadMilliVolts(pin);
    sum += buf[i];
    delayMicroseconds(100);
  }
  float mean = (float)sum / CT_SAMPLES;
  float sq = 0.0f;
  for (uint16_t i = 0; i < CT_SAMPLES; i++) {
    float d = (float)buf[i] - mean;
    sq += d * d;
  }
  float vrms = sqrtf(sq / CT_SAMPLES) / 1000.0f;   // volts
  return vrms * CT_AMPS_PER_VOLT;
}

void beep(uint8_t times) {
  for (uint8_t i = 0; i < times; i++) {
    digitalWrite(BUZZER_PIN, HIGH); delay(150);
    digitalWrite(BUZZER_PIN, LOW);  delay(150);
  }
}

void senseRoom(uint8_t r, uint32_t now) {
  RoomState &s = state[r];
  const RoomPins &p = ROOMS[r];

  if (digitalRead(p.pirPin) == HIGH) {
    s.lastMotionMs = now;
    s.everMotion = true;
  }
  s.occupied = s.everMotion && (now - s.lastMotionMs) < OCCUPANCY_TIMEOUT_MS;
  s.lightOn  = analogReadMilliVolts(p.ldrPin) > LIGHT_ON_MV;
  s.acAmps   = readCtAmps(p.ctPin);
  s.acOn     = s.acAmps > AC_ON_AMPS;

  bool emptyAndOn = !s.occupied && (s.lightOn || s.acOn);
  if (emptyAndOn) {
    if (s.wasteStartMs == 0) s.wasteStartMs = now ? now : 1;
    bool nowWaste = (now - s.wasteStartMs) >= WASTE_GRACE_MS;
    if (nowWaste && !s.waste) {               // rising edge -> alert once
      s.wasteAlerts++;
      Serial.printf("ALERT,%lu,%s,%s%s left ON in empty room\n",
                    (unsigned long)(now / 1000), p.name,
                    s.lightOn ? "Lights" : "",
                    s.acOn ? (s.lightOn ? " + AC" : "AC") : "");
      beep(3);
    }
    s.waste = nowWaste;
  } else {
    s.wasteStartMs = 0;
    s.waste = false;
  }
}

void logRooms(uint32_t now) {
  // DATA,uptime_s,room,occupied,light_on,ac_on,ac_amps,waste
  for (uint8_t r = 0; r < NUM_ROOMS; r++) {
    const RoomState &s = state[r];
    Serial.printf("DATA,%lu,%s,%d,%d,%d,%.2f,%d\n", (unsigned long)(now / 1000), ROOMS[r].name,
                  s.occupied, s.lightOn, s.acOn, s.acAmps, s.waste);
  }
}

#if ENABLE_WIFI
void handleRoot() {
  String h = F("<!DOCTYPE html><html><head><meta charset='utf-8'><meta http-equiv='refresh' content='10'>"
               "<meta name='viewport' content='width=device-width,initial-scale=1'>"
               "<title>Occupancy & Energy Monitor</title><style>body{font-family:Arial;margin:20px}"
               "table{border-collapse:collapse}td,th{border:1px solid #ccc;padding:6px 10px}"
               ".w{background:#f8cbad;font-weight:bold}</style></head><body>"
               "<h2>Workspace Occupancy &amp; Energy Monitor</h2><table><tr><th>Room</th><th>Occupied</th>"
               "<th>Lights</th><th>AC (A)</th><th>Status</th><th>Alerts</th></tr>");
  for (uint8_t r = 0; r < NUM_ROOMS; r++) {
    const RoomState &s = state[r];
    h += s.waste ? "<tr class='w'>" : "<tr>";
    h += "<td>" + String(ROOMS[r].name) + "</td><td>" + (s.occupied ? "Yes" : "No") + "</td><td>" +
         (s.lightOn ? "ON" : "off") + "</td><td>" + String(s.acAmps, 2) + (s.acOn ? " ON" : " off") + "</td><td>" +
         (s.waste ? "ENERGY WASTE" : "OK") + "</td><td>" + String(s.wasteAlerts) + "</td></tr>";
  }
  h += F("</table><p>Auto-refreshes every 10 s. Raw data: <a href='/csv'>/csv</a></p></body></html>");
  server.send(200, "text/html", h);
}

void handleCsv() {
  String c = "room,occupied,light_on,ac_on,ac_amps,waste\n";
  for (uint8_t r = 0; r < NUM_ROOMS; r++) {
    const RoomState &s = state[r];
    c += String(ROOMS[r].name) + "," + (s.occupied ? "1" : "0") + "," + (s.lightOn ? "1" : "0") + "," +
         (s.acOn ? "1" : "0") + "," + String(s.acAmps, 2) + "," + (s.waste ? "1" : "0") + "\n";
  }
  server.send(200, "text/csv", c);
}
#endif

// --------------------------------- main ------------------------------------
void setup() {
  Serial.begin(115200);
  pinMode(BUZZER_PIN, OUTPUT);
  pinMode(STATUS_LED, OUTPUT);
  for (uint8_t r = 0; r < NUM_ROOMS; r++) {
    pinMode(ROOMS[r].pirPin, INPUT);
    analogSetPinAttenuation(ROOMS[r].ldrPin, ADC_11db);   // full 0-3.3 V range
    analogSetPinAttenuation(ROOMS[r].ctPin, ADC_11db);
  }
  delay(500);
  Serial.println("#Workspace Occupancy and Energy Monitor started");
  Serial.println("#HEADER,uptime_s,room,occupied,light_on,ac_on,ac_amps,waste");

#if ENABLE_WIFI
  WiFi.mode(WIFI_STA);
  WiFi.begin(WIFI_SSID, WIFI_PASS);
  uint32_t t0 = millis();
  while (WiFi.status() != WL_CONNECTED && millis() - t0 < 15000) delay(250);
  if (WiFi.status() == WL_CONNECTED) {
    Serial.print("#Web page: http://");
    Serial.println(WiFi.localIP());
    server.on("/", handleRoot);
    server.on("/csv", handleCsv);
    server.begin();
  } else {
    Serial.println("#Wi-Fi not connected, continuing with serial logging only");
  }
#endif
  // PIR modules need ~30-60 s to settle after power-up
  Serial.println("#Warming up PIR sensors (30 s)...");
  delay(30000);
  lastLogMs = millis();
}

void loop() {
  uint32_t now = millis();

  if (now - lastSenseMs >= SENSE_INTERVAL_MS) {
    lastSenseMs = now;
    bool anyWaste = false;
    for (uint8_t r = 0; r < NUM_ROOMS; r++) {
      senseRoom(r, now);
      anyWaste |= state[r].waste;
    }
    digitalWrite(STATUS_LED, anyWaste ? HIGH : LOW);   // LED stays on while any room is wasting
  }

  if (now - lastLogMs >= LOG_INTERVAL_MS) {
    lastLogMs = now;
    logRooms(now);
  }

#if ENABLE_WIFI
  server.handleClient();
#endif
}
