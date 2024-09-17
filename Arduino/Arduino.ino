#define pulPin 2
#define enablePin 3
#define dirPin 4
#define comHigh 5
#define homeProxPinLow 6
#define homeProxPinHigh 7
#define homeProxPin 8
#define endProxPinLow 9
#define endProxPinHigh 10
#define endProxPin 11

String msg;
void processData(String capturedString);
void pulse(void);
void enableStepper(void);
void stepperClockwise(void);
void disableStepper(void);
void stepperAntiClockwise(void);
String split(String data, char separator, int index);
bool home = 0;
bool end = 0;
bool locate = 0;
bool stringComplete = false;

int speed;
float move = 0;
int steps = 0;
String inputString = "";
bool stopStatus = true;
bool direction = 1;

void serialEvent() {
  while(Serial.available()) {
    char inChar = (char)Serial.read();
    if (inChar == '\n') {
      stringComplete = true;
    }
    else {
      inputString += inChar;
    }
  }
}

String split(String data, char separator, int index){
  int found = 0;
  int strIndex[] = {0, -1};
  int maxIndex = data.length() - 1;
  for (int i = 0; i <= maxIndex && found <= index; i++){
    if (data.charAt(i) == separator || i == maxIndex){
      found++;
      strIndex[0] = strIndex[1] + 1;
      strIndex[1] = (i == maxIndex) ? i+1 : i;
    }
  }
  return found > index ? data.substring(strIndex[0], strIndex[1]) : "";
}

void pulse(void) {
  if ((digitalRead(homeProxPin) == HIGH && direction == 0) || (digitalRead(endProxPin) == HIGH && direction == 1)) {
    digitalWrite(pulPin, HIGH);
    delayMicroseconds(speed);
    digitalWrite(pulPin, LOW);
    delayMicroseconds(speed);
    Serial.flush();
    Serial.println("p");
  }
  else if (digitalRead(homeProxPin) == LOW) {
    Serial.flush();
    Serial.println("h");
  }
  else if (digitalRead(endProxPin) == LOW) {
    Serial.flush();
    Serial.println("n");
  }
}


void enableStepper(void) {
  digitalWrite(enablePin, HIGH);
  Serial.flush();
  Serial.println("e");
}

void disableStepper(void) {
  digitalWrite(enablePin, LOW);
  Serial.flush();
  Serial.println("d");
}

void stepperClockwise(void) {
  digitalWrite(dirPin, HIGH);
  direction = 1;
  Serial.flush();
  Serial.println("c");
}

void stepperAntiClockwise(void) {
  digitalWrite(dirPin, LOW);
  direction = 0;
  Serial.flush();
  Serial.println("a");
}

void setup() {
  Serial.begin(9600);
  Serial.setTimeout(1);
  inputString.reserve(256);

  pinMode(pulPin,OUTPUT);
  pinMode(enablePin,OUTPUT);
  pinMode(dirPin,OUTPUT);
  pinMode(comHigh,OUTPUT);
  pinMode(homeProxPinLow,OUTPUT);
  pinMode(homeProxPinHigh,OUTPUT);
  pinMode(homeProxPin,INPUT);
  pinMode(endProxPinLow,OUTPUT);
  pinMode(endProxPinHigh,OUTPUT);
  pinMode(endProxPin,INPUT);

  digitalWrite(comHigh, HIGH);
  digitalWrite(dirPin, HIGH);
  digitalWrite(enablePin, HIGH);
  digitalWrite(homeProxPinHigh, HIGH);
  digitalWrite(endProxPinHigh, HIGH);
  digitalWrite(endProxPinLow, LOW);
  digitalWrite(endProxPinLow, LOW);
}

void loop() {
  if (stringComplete) {
    if (inputString == "P") {
      pulse();
    }
    else if (inputString == "E") {
        enableStepper();
    }
    else if (inputString == "C") {
        stepperClockwise(); 
    }
    else if (inputString == "A") {
        stepperAntiClockwise();
    }
    else if (inputString == "D") {
      disableStepper();
    }
    else if (inputString == "H") {
      home = 1;
      stopStatus = false;
    }
    else if (inputString == "N") {
      end = 1;
      stopStatus = false;
    }
    else if (inputString == "S") {
      stopStatus = true;
    }
    else if (inputString == "L") {
      locate = 1;
      stopStatus = false;
      steps = 0;
    }
    else{
      String key = split(inputString, ':', 0);
      if (key == "pulse") {
        String value = split(inputString, ':', 1);
        speed = value.toInt();
      }
      else if (key == "move") {
        String value = split(inputString, ':', 1);
        move = value.toInt();
        stopStatus = false;
      }
    }
    inputString = "";
    stringComplete = false;
  }
  if (direction == 0 && home == 1) {
    if (digitalRead(homeProxPin) == HIGH && stopStatus == false) {
      digitalWrite(pulPin, HIGH);
      delayMicroseconds(speed);
      digitalWrite(pulPin, LOW);
      delayMicroseconds(speed);
    }
    else if (digitalRead(homeProxPin) == LOW) {
      Serial.flush();
      Serial.println("h");
      home = 0;
    }
    else if (stopStatus == true) {
      Serial.flush();
      Serial.println("s");
      home = 0;
    }
  }
  if (direction == 1 && end == 1) {
    if (digitalRead(endProxPin) == HIGH && stopStatus == false) {
      digitalWrite(pulPin, HIGH);
      delayMicroseconds(speed);
      digitalWrite(pulPin, LOW);
      delayMicroseconds(speed);
    }
    else if (digitalRead(endProxPin) == LOW) {
      Serial.flush();
      Serial.println("n");
      end = 0;
    }
    else if (stopStatus == true) {
      Serial.flush();
      Serial.println("s");
      end = 0;
    }
  }
  if (move > 0) {
    if ((digitalRead(homeProxPin) == HIGH || direction == 1) && (digitalRead(endProxPin) == HIGH || direction == 0) && stopStatus == false) {
      digitalWrite(pulPin, HIGH);
      delayMicroseconds(speed);
      digitalWrite(pulPin, LOW);
      delayMicroseconds(speed);
      move = move - 1;
      if (move < 1){
        Serial.flush();
        Serial.println("m");
      }
    }
    else if (digitalRead(homeProxPin) == LOW) {
      Serial.flush();
      Serial.println("h");
      move = 0;
    }
    else if (digitalRead(endProxPin) == LOW) {
      Serial.flush();
      Serial.println("e");
      move = 0;
    }
    else if (stopStatus == true) {
      Serial.flush();
      Serial.println("s");
      move = 0;
    }
  }
  if (direction == 0 && locate == 1) {
    if (digitalRead(homeProxPin) == HIGH && stopStatus == false) {
      digitalWrite(pulPin, HIGH);
      delayMicroseconds(speed);
      digitalWrite(pulPin, LOW);
      delayMicroseconds(speed);
      steps = steps + 1;
    }
    else if (digitalRead(homeProxPin) == LOW) {
      Serial.flush();
      Serial.println(steps);
      locate = 0;
      steps = 0;
    }
    else if (stopStatus == true) {
      Serial.flush();
      Serial.println("s");
      locate = 0;
      steps = 0;
    }
  }
}