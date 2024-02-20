import uno
import unohelper
from unidecode import unidecode
from com.sun.star.task import XJobExecutor
from com.sun.star.document import XDocumentEventListener



 
class Wavelet( unohelper.Base, XJobExecutor):
    def __init__( self, ctx ):
        self.ctx = ctx
    
    def trigger( self , args ):
        desktop = self.ctx.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", self.ctx)
        doc = desktop.getCurrentComponent()
        text = doc.Text
        doc.addDocumentEventListener(DocumentEventListener(doc))


# Starting from Python IDE
def main():
    try:
        ctx = XSCRIPTCONTEXT
    except NameError:
        ctx = officehelper.bootstrap()
        if ctx is None:
            print("ERROR: Could not bootstrap default Office.")
            sys.exit(1)

# Starting from command line
if __name__ == "__main__":
    main()

class DocumentEventListener(unohelper.Base, XDocumentEventListener):

    def __init__(self, doc):
        self._doc = doc

    def documentEventOccured(self, event):
        # ~ app.debug(event.EventName)
        if event.EventName == 'OnLayoutFinished':
            self._search()
        return

    def _search(self):
        current = self._doc.CurrentController.ViewCursor.Start
        cursor = self._doc.Text.createTextCursorByRange(current)
        cursor.gotoStartOfWord(False)
        cursor.gotoEndOfWord(True)

        search = self._doc.createSearchDescriptor()
        search.SearchWords = True
        search.SearchString = cursor.String

        found = self._doc.findAll(search)

        strict_pattern = self._build_strict_pattern(cursor.String)
        search.SearchRegularExpression = True
        search.SearchString = strict_pattern
        strict_found = self._doc.findAll(search)

        if found.Count > 1 or strict_found.Count > 1:
            cursor.CharBackColor = 0xFFFF00
        elif cursor.CharBackColor == 0xFFFF00:
            cursor.CharBackColor = -1
        return

    def _build_strict_pattern( self , word ):

        pattern = ''
        for char in word:
            if unidecode(char) == 'a':
                pattern += '[aáàâä]'
            elif unidecode(char) == 'e':
                pattern += '[eéèêë]'
            elif unidecode(char) == 'i':
                pattern += '[iíìîï]'
            elif unidecode(char) == 'o':
                pattern += '[oóòôö]'
            elif unidecode(char) == 'u':
                pattern += '[uúùûü]'
            else:
                pattern += char

        return r"\b(" + pattern + r")\b"


# pythonloader loads a static g_ImplementationHelper variable
g_ImplementationHelper = unohelper.ImplementationHelper()
g_ImplementationHelper.addImplementation(
    Wavelet,  # UNO object class
    "org.extension.sample.do",  # implementation name (customize for yourself)
    ("com.sun.star.task.Job",), )  # implemented services (only 1)

    