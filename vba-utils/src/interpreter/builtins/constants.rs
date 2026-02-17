use crate::context::Value;

pub(crate) fn resolve_builtin_identifier(name: &str) -> Option<Value> {
    match name {
        "vbCalGreg" => Some(Value::Integer( 0)),
        "vbCalHijri" => Some(Value::Integer( 1)),

        // CallType constants
        "vbMethod" => Some(Value::Integer( -1)),
        "vbGet" => Some(Value::Integer( -2)),
        "vbLet" => Some(Value::Integer( -4)),
        "vbSet" => Some(Value::Integer( -8)),

         // Color constants
        "vbBlack" => Some(Value::Integer( 0)),
        "vbRed" => Some(Value::Integer( 255)),
        "vbGreen" => Some(Value::Integer( 65280)),
        "vbYellow" => Some(Value::Integer( 65535)),
        "vbBlue" => Some(Value::Integer( 16711680)),
        "vbMagenta" => Some(Value::Integer( 16711935)),
        "vbCyan" => Some(Value::Integer( 16776960)),
        "vbWhite" => Some(Value::Integer( 16777215)),

         // Comparison constants
        "vbUseCompareOption"=> Some(Value::Integer(-1)),
        "vbBinaryCompare" => Some(Value::Integer( 0)),
        "vbTextCompare" => Some(Value::Integer( 1)),
        "vbDatabaseCompare" => Some(Value::Integer( 2)),

         // Day of Week constants
         "vbSunday" => Some(Value::Integer( 1)),
         "vbMonday" => Some(Value::Integer( 2)),
         "vbTuesday" => Some(Value::Integer( 3)),
         "vbWednesday" => Some(Value::Integer( 4)),
         "vbThursday" => Some(Value::Integer( 5)),
         "vbFriday" => Some(Value::Integer( 6)),
         "vbSaturday" => Some(Value::Integer( 7)),
         "vbUseSystemDayOfWeek" => Some(Value::Integer( 0)),

         // First Week of Year constants
         "vbUseSystem" => Some(Value::Integer( 0)),
         "vbFirstJan1" => Some(Value::Integer( 1)),
         "vbFirstFourDays" => Some(Value::Integer( 2)),
         "vbFirstFullWeek" => Some(Value::Integer( 3)),

         // Date/Time format constants
         "vbGeneralDate" => Some(Value::Integer( 0)),
         "vbLongDate" => Some(Value::Integer( 1)),
         "vbShortDate" => Some(Value::Integer( 2)),
         "vbLongTime" => Some(Value::Integer( 3)),
         "vbShortTime" => Some(Value::Integer( 4)),

         // Key Code Constants - Mouse Buttons
         "vbKeyLButton" => Some(Value::Integer( 1)),        // 0x1 - Left mouse button
         "vbKeyRButton" => Some(Value::Integer( 2)),        // 0x2 - Right mouse button
         "vbKeyCancel" => Some(Value::Integer( 3)),         // 0x3 - CANCEL key
         "vbKeyMButton" => Some(Value::Integer( 4)),        // 0x4 - Middle mouse button
         
         // Key Code Constants - Special Keys
         "vbKeyBack" => Some(Value::Integer( 8)),           // 0x8 - BACKSPACE key
         "vbKeyTab" => Some(Value::Integer( 9)),            // 0x9 - TAB key
         "vbKeyClear" => Some(Value::Integer( 12)),         // 0xC - CLEAR key
         "vbKeyReturn" => Some(Value::Integer( 13)),        // 0xD - ENTER key
         "vbKeyShift" => Some(Value::Integer( 16)),         // 0x10 - SHIFT key
         "vbKeyControl" => Some(Value::Integer( 17)),       // 0x11 - CTRL key
         "vbKeyMenu" => Some(Value::Integer( 18)),          // 0x12 - MENU key
         "vbKeyPause" => Some(Value::Integer( 19)),         // 0x13 - PAUSE key
         "vbKeyCapital" => Some(Value::Integer( 20)),       // 0x14 - CAPS LOCK key
         "vbKeyEscape" => Some(Value::Integer( 27)),        // 0x1B - ESC key
         "vbKeySpace" => Some(Value::Integer( 32)),         // 0x20 - SPACEBAR key
         
         // Key Code Constants - Navigation Keys
         "vbKeyPageUp" => Some(Value::Integer( 33)),        // 0x21 - PAGE UP key
         "vbKeyPageDown" => Some(Value::Integer( 34)),      // 0x22 - PAGE DOWN key
         "vbKeyEnd" => Some(Value::Integer( 35)),           // 0x23 - END key
         "vbKeyHome" => Some(Value::Integer( 36)),          // 0x24 - HOME key
         "vbKeyLeft" => Some(Value::Integer( 37)),          // 0x25 - LEFT ARROW key
         "vbKeyUp" => Some(Value::Integer( 38)),            // 0x26 - UP ARROW key
         "vbKeyRight" => Some(Value::Integer( 39)),         // 0x27 - RIGHT ARROW key
         "vbKeyDown" => Some(Value::Integer( 40)),          // 0x28 - DOWN ARROW key
         "vbKeySelect" => Some(Value::Integer( 41)),        // 0x29 - SELECT key
         "vbKeyPrint" => Some(Value::Integer( 42)),         // 0x2A - PRINT SCREEN key
         "vbKeyExecute" => Some(Value::Integer( 43)),       // 0x2B - EXECUTE key
         "vbKeySnapshot" => Some(Value::Integer( 44)),      // 0x2C - SNAPSHOT key
         "vbKeyInsert" => Some(Value::Integer( 45)),        // 0x2D - INSERT key
         "vbKeyDelete" => Some(Value::Integer( 46)),        // 0x2E - DELETE key
         "vbKeyHelp" => Some(Value::Integer( 47)),          // 0x2F - HELP key
         "vbKeyNumlock" => Some(Value::Integer( 144)),      // 0x90 - NUM LOCK key
         
         // Key Code Constants - Alphabetic Keys (A-Z)
         "vbKeyA" => Some(Value::Integer( 65)),             // ASCII 'A'
         "vbKeyB" => Some(Value::Integer( 66)),             // ASCII 'B'
         "vbKeyC" => Some(Value::Integer( 67)),             // ASCII 'C'
         "vbKeyD" => Some(Value::Integer( 68)),             // ASCII 'D'
         "vbKeyE" => Some(Value::Integer( 69)),             // ASCII 'E'
         "vbKeyF" => Some(Value::Integer( 70)),             // ASCII 'F'
         "vbKeyG" => Some(Value::Integer( 71)),             // ASCII 'G'
         "vbKeyH" => Some(Value::Integer( 72)),             // ASCII 'H'
         "vbKeyI" => Some(Value::Integer( 73)),             // ASCII 'I'
         "vbKeyJ" => Some(Value::Integer( 74)),             // ASCII 'J'
         "vbKeyK" => Some(Value::Integer( 75)),             // ASCII 'K'
         "vbKeyL" => Some(Value::Integer( 76)),             // ASCII 'L'
         "vbKeyM" => Some(Value::Integer( 77)),             // ASCII 'M'
         "vbKeyN" => Some(Value::Integer( 78)),             // ASCII 'N'
         "vbKeyO" => Some(Value::Integer( 79)),             // ASCII 'O'
         "vbKeyP" => Some(Value::Integer( 80)),             // ASCII 'P'
         "vbKeyQ" => Some(Value::Integer( 81)),             // ASCII 'Q'
         "vbKeyR" => Some(Value::Integer( 82)),             // ASCII 'R'
         "vbKeyS" => Some(Value::Integer( 83)),             // ASCII 'S'
         "vbKeyT" => Some(Value::Integer( 84)),             // ASCII 'T'
         "vbKeyU" => Some(Value::Integer( 85)),             // ASCII 'U'
         "vbKeyV" => Some(Value::Integer( 86)),             // ASCII 'V'
         "vbKeyW" => Some(Value::Integer( 87)),             // ASCII 'W'
         "vbKeyX" => Some(Value::Integer( 88)),             // ASCII 'X'
         "vbKeyY" => Some(Value::Integer( 89)),             // ASCII 'Y'
         "vbKeyZ" => Some(Value::Integer( 90)),             // ASCII 'Z'
         
         // Key Code Constants - Numeric Keys (0-9)
         "vbKey0" => Some(Value::Integer( 48)),             // ASCII '0'
         "vbKey1" => Some(Value::Integer( 49)),             // ASCII '1'
         "vbKey2" => Some(Value::Integer( 50)),             // ASCII '2'
         "vbKey3" => Some(Value::Integer( 51)),             // ASCII '3'
         "vbKey4" => Some(Value::Integer( 52)),             // ASCII '4'
         "vbKey5" => Some(Value::Integer( 53)),             // ASCII '5'
         "vbKey6" => Some(Value::Integer( 54)),             // ASCII '6'
         "vbKey7" => Some(Value::Integer( 55)),             // ASCII '7'
         "vbKey8" => Some(Value::Integer( 56)),             // ASCII '8'
         "vbKey9" => Some(Value::Integer( 57)),             // ASCII '9'
         
         // Key Code Constants - Numpad Keys
         "vbKeyNumpad0" => Some(Value::Integer( 96)),       // Numpad 0
         "vbKeyNumpad1" => Some(Value::Integer( 97)),       // Numpad 1
         "vbKeyNumpad2" => Some(Value::Integer( 98)),       // Numpad 2
         "vbKeyNumpad3" => Some(Value::Integer( 99)),       // Numpad 3
         "vbKeyNumpad4" => Some(Value::Integer( 100)),      // Numpad 4
         "vbKeyNumpad5" => Some(Value::Integer( 101)),      // Numpad 5
         "vbKeyNumpad6" => Some(Value::Integer( 102)),      // Numpad 6
         "vbKeyNumpad7" => Some(Value::Integer( 103)),      // Numpad 7
         "vbKeyNumpad8" => Some(Value::Integer( 104)),      // Numpad 8
         "vbKeyNumpad9" => Some(Value::Integer( 105)),      // Numpad 9
         "vbKeyMultiply" => Some(Value::Integer( 106)),     // Numpad * (multiply)
         "vbKeyAdd" => Some(Value::Integer( 107)),          // Numpad + (add)
         "vbKeySeparator" => Some(Value::Integer( 108)),    // Numpad separator
         "vbKeySubtract" => Some(Value::Integer( 109)),     // Numpad - (subtract)
         "vbKeyDecimal" => Some(Value::Integer( 110)),      // Numpad . (decimal)
         "vbKeyDivide" => Some(Value::Integer( 111)),       // Numpad / (divide)
         
         // Key Code Constants - Function Keys (F1-F16)
         "vbKeyF1" => Some(Value::Integer( 112)),           // F1 key
         "vbKeyF2" => Some(Value::Integer( 113)),           // F2 key
         "vbKeyF3" => Some(Value::Integer( 114)),           // F3 key
         "vbKeyF4" => Some(Value::Integer( 115)),           // F4 key
         "vbKeyF5" => Some(Value::Integer( 116)),           // F5 key
         "vbKeyF6" => Some(Value::Integer( 117)),           // F6 key
         "vbKeyF7" => Some(Value::Integer( 118)),           // F7 key
         "vbKeyF8" => Some(Value::Integer( 119)),           // F8 key
         "vbKeyF9" => Some(Value::Integer( 120)),           // F9 key
         "vbKeyF10" => Some(Value::Integer( 121)),          // F10 key
         "vbKeyF11" => Some(Value::Integer( 122)),          // F11 key
         "vbKeyF12" => Some(Value::Integer( 123)),          // F12 key
         "vbKeyF13" => Some(Value::Integer( 124)),          // F13 key
         "vbKeyF14" => Some(Value::Integer( 125)),          // F14 key
         "vbKeyF15" => Some(Value::Integer( 126)),          // F15 key
         "vbKeyF16" => Some(Value::Integer( 127)),          // F16 key

         
         // MsgBox constants
         "vbOKOnly" => Some(Value::Integer( 0)),
         "vbOKCancel" => Some(Value::Integer( 1)),
         "vbAbortRetryIgnore" => Some(Value::Integer( 2)),
         "vbYesNoCancel" => Some(Value::Integer( 3)),
         "vbYesNo" => Some(Value::Integer( 4)),
         "vbRetryCancel" => Some(Value::Integer( 5)),

         // MsgBox icon constants
         "vbCritical" => Some(Value::Integer( 16)),
         "vbQuestion" => Some(Value::Integer( 32)),
         "vbExclamation" => Some(Value::Integer( 48)),
         "vbInformation" => Some(Value::Integer( 64)),

         // MsgBox return Some(Values
         "vbOK" => Some(Value::Integer( 1)),
         "vbCancel" => Some(Value::Integer( 2)),
         "vbAbort" => Some(Value::Integer( 3)),
         "vbRetry" => Some(Value::Integer( 4)),
         "vbIgnore" => Some(Value::Integer( 5)),
         "vbYes" => Some(Value::Integer( 6)),
         "vbNo" => Some(Value::Integer( 7)),

          // Text case constants
         "vbUpperCase"   => Some(Value::Integer( 1)),
         "vbLowerCase"   => Some(Value::Integer( 2)),
         "vbProperCase"  => Some(Value::Integer( 3)),

         // String width and script constants
         "vbWide"        => Some(Value::Integer( 4)),
         "vbNarrow"      => Some(Value::Integer( 8)),
         "vbKatakana"    => Some(Value::Integer( 16)),
         "vbHiragana"    => Some(Value::Integer( 32)),

         // Unicode constants
         "vbUnicode"     => Some(Value::Integer( -1)),
         "vbFromUnicode" => Some(Value::Integer( -2)),

         "vbTrue" => Some(Value::Integer( -1)),
         "vbFalse" => Some(Value::Integer( 0)),
         "vbUseDefault" => Some(Value::Integer( 2)),

        

         "vbEmpty" => Some(Value::Integer( 0)),
         "vbNull" => Some(Value::Integer( 1)),
         "vbInteger" => Some(Value::Integer( 2)),
         "vbLong" => Some(Value::Integer( 3)),
         "vbSingle" => Some(Value::Integer( 4)),
         "vbDouble" => Some(Value::Integer( 5)),
         "vbCurrency" => Some(Value::Integer( 6)),
         "vbDate" => Some(Value::Integer( 7)),
         "vbString" => Some(Value::Integer( 8)),
         "vbObject" => Some(Value::Integer( 9)),
         "vbError" => Some(Value::Integer( 10)),
         "vbBoolean" => Some(Value::Integer( 11)),
         "vbVariant" => Some(Value::Integer( 12)),
         "vbDataObject" => Some(Value::Integer( 13)),
         "vbDecimal" => Some(Value::Integer( 14)),
         "vbByte" => Some(Value::Integer( 17)),
         "vbUserDefinedType" => Some(Value::Integer( 36)),
         "vbArray" => Some(Value::Integer( 8192)),



         "vbCrLf"       => Some(Value::String( "\r\n".to_string())),
         "vbCr"         => Some(Value::String( "\r".to_string())),
         "vbLf"         => Some(Value::String( "\n".to_string())),
         "vbNewLine"    => Some(Value::String( "\n".to_string())),       // same as vbLf in many contexts
         "vbNullChar"   => Some(Value::String( '\0'.to_string())),       // null character
         "vbNullString" => Some(Value::String( "".to_string())),         // empty string
         "vbObjectError"=> Some(Value::Integer( -2147221504)), // typical COM error base (example Some(Value)
         "vbTab"        => Some(Value::String( "\t".to_string())),
         "vbBack"       => Some(Value::String( '\x08'.to_string())),     // backspace character
         "vbFormFeed"   => Some(Value::String( '\x0C'.to_string())),     // form feed character
         "vbVerticalTab"=> Some(Value::String( '\x0B'.to_string())),     // vertical tab character


        _ => {
            //println!("⚠️ Unknown builtin constant: {}", name);
            return None;
        }
    }
}