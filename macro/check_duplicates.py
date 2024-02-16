from unidecode import unidecode
 
def check_duplicates():
    desktop = XSCRIPTCONTEXT.getDesktop()
    doc = desktop.getCurrentComponent()
    text = doc.Text
    try:
        cursor = text.createTextCursor()
        cursor.gotoStart(False)
        dictionary = {}

        while cursor.gotoNextWord(True):

            cursor.gotoPreviousWord(False)
            cursor.gotoEndOfWord(True)

            currentWord = unidecode(cursor.getString().strip().lower())
                
            if currentWord in dictionary:
                cursor.CharBackColor = 0xFFFF00
            else:
                dictionary[currentWord] = True
                cursor.CharBackColor = -1
                
            cursor.gotoPreviousWord(False)
            cursor.gotoNextWord(False)

    except Exception as e:
            print("Error:", e)
            pass


g_exportedScripts = check_duplicates,
