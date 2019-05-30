/**
*  Mock up software
*/
const int sensorPin = A0;   // Analog input pin that senses Vout
int sensorValue = 0;        // sensorPin default value
float Vin       = 3.3;      // Input voltage
float Vout      = 0;        // Vout default value
float Rref      = 353.9;    // Reference resistor's value in ohms
float y         = 0;
float R         = 0;
int i           = 0;
int inByte      = 0;
char str_temp[6];

void setup() {
  Serial.begin(9600);
  Serial.println("setup");
} // setup

void loop() {
  while (Serial.available() > 0) {
    int val = Serial.read();
    if (val == '\n') {
      if (inByte == 'M') {

        //
        // Mock up data
        //
        //float temp = (random(1, 2000) / 10.0) + 250;
        //dtostrf(temp, 4, 2, str_temp);

        //
        // Real data
        //
        y=0;
        for (i = 0;i<20;i++) {
          sensorValue = analogRead(sensorPin);
          Vout = (Vin * sensorValue) / 1023;    // Convert Vout to volts
          R = Rref * (1 / ((Vin / Vout) - 1));
          y=y+R;
          delay(40);
        }
        y=y/20;
        dtostrf(y, 4, 2, str_temp);

        //
        // Send a message
        //
        Serial.write(str_temp);
        Serial.write('\n');
        continue;
      }
      continue;
    } // '\n'

    inByte = val;
  } // Serial.available
} // loop
