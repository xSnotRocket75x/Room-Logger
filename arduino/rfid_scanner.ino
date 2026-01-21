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

// Helper function to print a single hex digit safely with delay
void printHexDigit(uint8_t digit) {
  // Convert value 0-15 to character 0-9 or A-F
  if (digit < 10) {
    Keyboard.print((char)('0' + digit));
  } else {
    Keyboard.print((char)('A' + (digit - 10)));
  }
  // CRITICAL: Small delay to let Linux process the key state (Shift Up/Down)
  delay(10); 
}

void loop(void) {
  uint8_t success;
  uint8_t uid[] = { 0, 0, 0, 0, 0, 0, 0 };
  uint8_t uidLength;

  success = nfc.readPassiveTargetID(PN532_MIFARE_ISO14443A, uid, &uidLength);

  if (success) {
    for (uint8_t i = 0; i < uidLength; i++) {
      // Print the high nibble (first digit)
      printHexDigit(uid[i] >> 4);
      
      // Print the low nibble (second digit)
      printHexDigit(uid[i] & 0x0F);
    }

    Keyboard.write(KEY_RETURN);

    // Visual feedback
    digitalWrite(LED_BUILTIN_TX, LOW); // TXLED1 equivalent usually
    delay(500);
    digitalWrite(LED_BUILTIN_TX, HIGH);

    delay(2000);
  }
}