try:
    import keyboard
    import time
except ImportError:
    keyboard = None

KEY_NAMES = ['BACKSPACE','TAB','CLEAR','ENTER','SHIFT','CTRL','ALT','PAUSE','CAPS LOCK','CONTROL-BREAK PROCESSING',
            'IME KANA MODE','IME HANGUEL MODE','IME HANGUL MODE','IME JUNJA MODE','IME FINAL MODE','IME HANJA MODE',
            'IME KANJI MODE','ESC','IME CONVERT','IME NONCONVERT','IME ACCEPT','IME MODE CHANGE REQUEST','SPACEBAR',
            'PAGE UP','PAGE DOWN','END','HOME','LEFT','UP','RIGHT','DOWN','SELECT','PRINT','EXECUTE','PRINT SCREEN',
            'INSERT','DELETE','HELP','0','1','2','3','4','5','6','7','8','9','A','B','C','D','E','F','G','H','I','J',
            'K','L','M','N','O','P','Q','R','S','T','U','V','W','X','Y','Z','LEFT WINDOWS','RIGHT WINDOWS','APPLICATIONS',
            'SLEEP','0','1','2','3','4','5','6','7','8','9','*','+','SEPARATOR','-','DECIMAL','/','F1','F2','F3','F4','F5',
            'F6','F7','F8','F9','F10','F11','F12','F13','F14','F15','F16','F17','F18','F19','F20','F21','F22','F23','F24',
            'NUM LOCK','SCROLL LOCK','LEFT SHIFT','RIGHT SHIFT','LEFT CTRL','RIGHT CTRL','LEFT MENU','RIGHT MENU',
            'BROWSER BACK','BROWSER FORWARD','BROWSER REFRESH','BROWSER STOP','BROWSER SEARCH KEY','BROWSER FAVORITES',
            'BROWSER START AND HOME','VOLUME MUTE','VOLUME DOWN','VOLUME UP','NEXT TRACK','PREVIOUS TRACK','STOP MEDIA',
            'PLAY/PAUSE MEDIA','START MAIL','SELECT MEDIA','START APPLICATION 1','START APPLICATION 2','+',',','-','.',
            'IME PROCESS','ATTN','CRSEL','EXSEL','ERASE EOF','PLAY','ZOOM','RESERVED ','PA1','CLEAR']

class Keyboard:

    """
    Library that simulates keyboard input.

    Available keys are listed below:

        ['BACKSPACE','TAB','CLEAR','ENTER','SHIFT','CTRL','ALT','PAUSE','CAPS LOCK','CONTROL-BREAK PROCESSING',
        'IME KANA MODE','IME HANGUEL MODE','IME HANGUL MODE','IME JUNJA MODE','IME FINAL MODE','IME HANJA MODE',
        'IME KANJI MODE','ESC','IME CONVERT','IME NONCONVERT','IME ACCEPT','IME MODE CHANGE REQUEST','SPACEBAR',
        'PAGE UP','PAGE DOWN','END','HOME','LEFT','UP','RIGHT','DOWN','SELECT','PRINT','EXECUTE','PRINT SCREEN',
        'INSERT','DELETE','HELP','0','1','2','3','4','5','6','7','8','9','A','B','C','D','E','F','G','H','I','J',
        'K','L','M','N','O','P','Q','R','S','T','U','V','W','X','Y','Z','LEFT WINDOWS','RIGHT WINDOWS','APPLICATIONS',
        'SLEEP','0','1','2','3','4','5','6','7','8','9','*','+','SEPARATOR','-','DECIMAL','/','F1','F2','F3','F4','F5',
        'F6','F7','F8','F9','F10','F11','F12','F13','F14','F15','F16','F17','F18','F19','F20','F21','F22','F23','F24',
        'NUM LOCK','SCROLL LOCK','LEFT SHIFT','RIGHT SHIFT','LEFT CTRL','RIGHT CTRL','LEFT MENU','RIGHT MENU',
        'BROWSER BACK','BROWSER FORWARD','BROWSER REFRESH','BROWSER STOP','BROWSER SEARCH KEY','BROWSER FAVORITES',
        'BROWSER START AND HOME','VOLUME MUTE','VOLUME DOWN','VOLUME UP','NEXT TRACK','PREVIOUS TRACK','STOP MEDIA',
        'PLAY/PAUSE MEDIA','START MAIL','SELECT MEDIA','START APPLICATION 1','START APPLICATION 2','+',',','-','.',
        'IME PROCESS','ATTN','CRSEL','EXSEL','ERASE EOF','PLAY','ZOOM','RESERVED ','PA1','CLEAR']
    """
    def __init__(self):
        print("KeyboardLibrary library initialised")

    def native_type(self, key):

        """
            Sends key to the current window. Use {SPACE}, {TAB}, {ENTER} for spaces, tabs and new lines and should be given in capital letters.

            |***TestCases*** |
            1. Enter Text
            Native Type | Hello World

            2. Enter Text using {SPACE}
            Native Type | Hello{SPACE}World

            3. Enter Text using {SPACE} and press Enter button at the end.
            Native Type | Hello{SPACE}World{ENTER}

            4. Enter Text using {TAB} and press Enter button at the end.
            Native Type | Hello{TAB}World{ENTER}

        """
        if keyboard is None:
            raise AssertionError("This keyword is supported on Windows platform only.")
        if not '{SPACE}' and '{TAB}' and '{ENTER}' in key:
            keyboard.write(key)
        string = key
        if '{SPACE}' in key:
            string = str(key).replace("{SPACE}", ' ')
        if '{TAB}' in string:
            string = str(string).replace("{TAB}", '    ')
        try:
            if '{ENTER}' in string:
                keyboard.write(str(string).replace("{ENTER}", ''))
                keyboard.send('enter')
            else:
                keyboard.write(str(string))
        except ValueError as e:
            raise AssertionError("Invalid keyboard key found : {}. Use {SPACE}, {TAB}, {ENTER} for spaces, tabs and new lines ".format(e))

    def select_all(self):
        """
        Simulates "Ctrl+A" key combination to select all elements.
        """
        if keyboard is None:
            raise AssertionError("This keyword is supported on Windows platform only.")
        keyboard.send('ctrl+A')

    def _convert_to_valid_special_key(self, key):

        key = str(key).lower()
        if key.startswith('key.'):
            key = key.split('key.', 1)[1]
            return key
        elif len(key) > 1:
            return None

    def _validate_keys(self, keys):
        valid_keys = []
        for key in keys:
            valid_key = self._convert_to_valid_special_key(key)
            valid_keys.append(valid_key)
        return valid_keys

    def press_combination(self, *keys):
        '''Press given keyboard keys.

        All keyboard keys must be prefixed with ``Key.``.

        KeyboardLibrary keys are case-insensitive:

        Available keys are listed below:

        ['BACKSPACE','TAB','CLEAR','ENTER','SHIFT','CTRL','ALT','PAUSE','CAPS LOCK',CONTROL-BREAK PROCESSING',
        'IME KANA MODE','IME HANGUEL MODE','IME HANGUL MODE','IME JUNJA MODE','IME FINAL MODE','IME HANJA MODE',
        'IME KANJI MODE','ESC','IME CONVERT','IME NONCONVERT','IME ACCEPT','IME MODE CHANGE REQUEST','SPACEBAR',
        'PAGE UP','PAGE DOWN','END','HOME','LEFT','UP','RIGHT','DOWN','SELECT','PRINT','EXECUTE','PRINT SCREEN',
        'INSERT','DELETE','HELP','0','1','2','3','4','5','6','7','8','9','A','B','C','D','E','F','G','H','I','J',
        'K','L','M','N','O','P','Q','R','S','T','U','V','W','X','Y','Z','LEFT WINDOWS','RIGHT WINDOWS','APPLICATIONS',
        'SLEEP','0','1','2','3','4','5','6','7','8','9','*','+','SEPARATOR','-','DECIMAL','/','F1','F2','F3','F4','F5',
        'F6','F7','F8','F9','F10','F11','F12','F13','F14','F15','F16','F17','F18','F19','F20','F21','F22','F23','F24',
        'NUM LOCK','SCROLL LOCK','LEFT SHIFT','RIGHT SHIFT','LEFT CTRL','RIGHT CTRL','LEFT MENU','RIGHT MENU',
        'BROWSER BACK','BROWSER FORWARD','BROWSER REFRESH','BROWSER STOP','BROWSER SEARCH KEY','BROWSER FAVORITES',
        'BROWSER START AND HOME','VOLUME MUTE','VOLUME DOWN','VOLUME UP','NEXT TRACK','PREVIOUS TRACK','STOP MEDIA',
        'PLAY/PAUSE MEDIA','START MAIL','SELECT MEDIA','START APPLICATION 1','START APPLICATION 2','+',',','-','.',
        'IME PROCESS','ATTN','CRSEL','EXSEL','ERASE EOF','PLAY','ZOOM','RESERVED ','PA1','CLEAR']

        | Press Combination | KEY.ENTER |
        | Press Combination | KEY.CTRL | KEY.C |
        | Press Combination | KEY.CTRL | KEY.V | 
        | Press Combination | KEY.END |

        [https://pyautogui.readthedocs.org/en/latest/keyboard.html#keyboard-keys|
        See valid keyboard keys here].
        '''
        keys = self._validate_keys(keys)
        keystopress = "+".join(str(i) for i in keys)
        try:
            keyboard.send(keystopress)
        except ValueError as e:
            invalid_key = str(e).split()[1]
            raise AssertionError("Invalid keyboard key found :{}. Please provide a valid key from the below list of keys:\n{}".format(invalid_key.upper(), KEY_NAMES))