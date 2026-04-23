#include <Servo.h>
#include <Wire.h>
#include <Adafruit_GFX.h>
#include <Adafruit_SSD1306.h>
#include <Adafruit_INA219.h>

Adafruit_INA219 ina219;  
Adafruit_SSD1306 display(128, 64, &Wire);


const int Lb = A8; // ldr linksboven
const int Lo = A9; // ldr linksonder
const int Rb = A10; // ldr rechtsboven
const int Ro = A11; // ldr rechtsonder

Servo servoX;
Servo servoY;

int posX = random(89, 92);  // 89 t/m 91
int posY = random(89, 92);  // 89 t/m 91


unsigned long lastMove = 0; // servo's niet constant laten bewegen 
const int moveDelay = 60; // maakt de servo's smooth

const int minX = 20; // max en minimum van servo
const int maxX = 160;

const int minY = 20;
const int maxY = 160;


const int buttonDisplay = 4;   // wisselt tussen LDR/paneel
const int buttonMeasure = 5;   // start 10s meting

bool showLDR = false;          // wisselt tussen LDR/paneel, true = ldr
bool measuring = false;        // true = meting is bezig 
bool measurementDone = false;  // true = meting is klaar 

unsigned long measureStart = 0; // meting begind
float sumPower = 0; // optelsom van vermogen 
int samples = 0; //aantal metingen

bool lastButtonDisplay = HIGH;   // onthoudt vorige staat van display-knop
bool lastButtonMeasure = HIGH;   // onthoudt vorige staat van meet-knop 

// smoothing
int smoothRead(int pin) {
  long total = 0;
  for (int i = 0; i < 5; i++) { // sneller maken
    total += analogRead(pin);
    delayMicroseconds(800);
  }
  return total / 5;
}

void setup() {
  Serial.begin(9600);
  servoX.attach(2); // D2
  servoY.attach(3); // D3
  servoX.write(posX); // zet servo in 90 grade
  servoY.write(posY); // zet servo in 90 grade
 
  ina219.begin();
 
  display.begin(SSD1306_SWITCHCAPVCC, 0x3C);
  display.clearDisplay();
  display.setTextSize(1);
  display.setTextColor(SSD1306_WHITE);
  display.setCursor(10, 0);
  display.println("Solar");
  display.setCursor(10, 20);
  display.println("Tracker");

  for (int i = 0; i <= 120; i++) {
    display.drawRect(4, 50, 120, 10, SSD1306_WHITE);
    display.fillRect(4, 50, i, 10, SSD1306_WHITE);
    display.display();
    delay(10);
  }
   
  pinMode(buttonDisplay, INPUT_PULLUP); // zegt dat het een pull‑up weerstand is
  pinMode(buttonMeasure, INPUT_PULLUP); // zegt dat het een pull‑up weerstand is
    
}

void loop() {
// dispay switch
bool readDisplay = digitalRead(buttonDisplay);
if (readDisplay == LOW && lastButtonDisplay == HIGH){

  if(measurementDone) {
    measurementDone = false; // zorgen dat we niet vastzitten in de meettoesland
    samples = 0;
  }
  else if(!measuring){
    showLDR = !showLDR; // switch tussen ldr en zonnepaneel
  }
}
lastButtonDisplay = readDisplay;

// 10s meting 
bool readMeasure = digitalRead(buttonMeasure); // huidige staat knop

if (readMeasure == LOW && lastButtonMeasure == HIGH && !measuring) {
  measuring = true;
  measurementDone = false;
  measureStart = millis();
  sumPower = 0;   // begin leeg
  samples = 0;
}
lastButtonMeasure = readMeasure;

//LDR Tracking
int LB = smoothRead(Lb);
int LO = smoothRead(Lo);
int RB = smoothRead(Rb);
int RO = smoothRead(Ro);

// INA219 meting
float busVoltage = ina219.getBusVoltage_V();
float current_mA = ina219.getCurrent_mA();
float power_mW   = ina219.getPower_mW();

//10s meting
if (measuring) {
  sumPower += power_mW;
  samples++;

  if (millis() - measureStart >= 10000){
    measuring = false;
    measurementDone = true;
  }
}

//scherm
display.clearDisplay();

if (measuring){
  // 10s meting
  display.setCursor(0, 0);
  display.print("Meting bezig...");
  display.setCursor(0, 20);
  display.print((millis() - measureStart) / 1000); // tijdweergave 
  display.print(" / 10 s");
  display.display(); // versturing naar scherm
}
else if (measurementDone){
  //toon gemiddelde
  float avgPower = sumPower / samples; // mW uitrekenen 
  display.setCursor(0, 0);
  display.print("Gemiddelde 10s:");
  display.setCursor(0, 20);
  display.print(avgPower, 1);
  display.print(" mW");
  display.display();
}
else {
// normale mode 
  if(showLDR){
    display.setCursor(0, 0);
    display.print("LDR waarden:");
    display.setCursor(0, 16); display.print("LB: "); display.print(LB);
    display.setCursor(0, 28); display.print("LO: "); display.print(LO);
    display.setCursor(0, 40); display.print("RB: "); display.print(RB);
    display.setCursor(0, 52); display.print("RO: "); display.print(RO);
  } else {
    display.setCursor(0, 0);
    display.print("Paneel:");
    display.setCursor(0, 16); display.print("U: "); display.print(busVoltage, 2);
    display.setCursor(0, 32); display.print("I: "); display.print(current_mA, 1);
    display.setCursor(0, 48); display.print("P: "); display.print(power_mW, 1);
  }
  display.display(); // sturen naar display 
}

//tracking 
int topAvg    = (LB + RB) / 2;
int bottomAvg = (LO + RO) / 2;
int leftAvg   = (LB + LO) / 2;
int rightAvg  = (RB + RO) / 2;

int diffY = topAvg - bottomAvg;
int diffX = rightAvg - leftAvg;

const int threshold = 50; //maakt het smooth

// dynamische stapgrootte
int stepX = constrain(abs(diffX) / 20, 1, 4);
int stepY = constrain(abs(diffY) / 20, 1, 4);

//tracking
if (millis() - lastMove > moveDelay) {

  if (abs(diffX) > threshold) {
    posX += (diffX > 0 ? stepX : -stepX);
  }

  if (abs(diffY) > threshold) {
    posY += (diffY > 0 ? -stepY : +stepY); // jouw omgedraaide Y
  }

  posX = constrain(posX, minX, maxX);
  posY = constrain(posY, minY, maxY);

  servoX.write(posX);
  servoY.write(posY);

  lastMove = millis();
}

}