/* Mock up software
*/
int inByte = 0;
char str_temp[6];

void setup() {
  Serial.begin(9600);
}

void loop() {
  while (Serial.available() > 0) {
    int val = Serial.read();
    if (val == '\n') {
      if (inByte == 'M') {
        float temp = (random(1, 2000) / 10.0) + 250;
        dtostrf(temp, 4, 2, str_temp);
        Serial.write(str_temp);
        Serial.write('\n');
        return;
      }
//      Serial.println(inByte);
      return;
    }

    inByte = val;
  }
}
