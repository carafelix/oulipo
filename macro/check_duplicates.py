from unidecode import unidecode
 
def check_duplicates():
    desktop = XSCRIPTCONTEXT.getDesktop()
    doc = desktop.getCurrentComponent()
    text = doc.Text
    try:
        cursor = text.createTextCursor()
        cursor.gotoStart(False)
        dictionary = {}

        while cursor.goRight(1, False):
            cursor.goLeft(1, False)
            cursor.gotoEndOfWord(True)

            currentWord = unidecode(cursor.getString().strip().lower())
            print(currentWord)
                
            if currentWord in dictionary:
                cursor.CharBackColor = 0xFFFF00
            else:
                dictionary[currentWord] = True
                cursor.CharBackColor = -1
            
            cursor.gotoNextWord(False)

    except Exception as e:
            print("Error:", e)
            pass


g_exportedScripts = check_duplicates,
