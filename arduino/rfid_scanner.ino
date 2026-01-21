#include <Wire.h>
#include <Adafruit_PN532.h>
#include <Keyboard.h>

#define PN532_IRQ (6)
#define PN532_RESET (7)

Adafruit_PN532 nfc(PN532_IRQ, PN532_RESET);

void setup(void) {
  nfc.begin();
  uint32_t versiondata = nfc.getFirmwareVersion();

  if (!versiondata) {
    while (1) {
      delay(1000);
    }
  }
  
  nfc.SAMConfig();
  Keyboard.begin();
}

void loop(void) {
  uint8_t success;
  uint8_t uid[] = { 0, 0, 0, 0, 0, 0, 0 };
  uint8_t uidLength;

  success = nfc.readPassiveTargetID(PN532_MIFARE_ISO14443A, uid, &uidLength);

  if (success) {
    for (uint8_t i = 0; i < uidLength; i++) {
      if (uid[i] < 0x10) {
        Keyboard.print("0");
      }
      Keyboard.print(uid[i], HEX);
    }

    Keyboard.write(KEY_RETURN);

    TXLED1;
    delay(500);
    TXLED0;

    delay(2000);
  }
}