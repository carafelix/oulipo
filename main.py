import uno
import unohelper
from com.sun.star.task import XJobExecutor
from com.sun.star.text import XTextDocument
from com.sun.star.lang import XEventListener

 
class Wavelet( unohelper.Base, XJobExecutor, XEventListener ):
    def __init__( self, ctx ):
        self.ctx = ctx
    
    def trigger( self ):
        desktop = self.ctx.ServiceManager.createInstanceWithContext("com.sun.star.frame.Desktop", self.ctx)
        doc = desktop.getCurrentComponent()
        text = doc.Text
        try:
            cursor = text.createTextCursor()
            cursor.gotoStart(False)
            dictionary = {}

            while cursor.gotoNextWord(True):

                cursor.gotoPreviousWord(False)
                cursor.gotoEndOfWord(True)

                currentWord = cursor.getString().strip().lower()
                
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


# Starting from Python IDE
def main():
    try:
        ctx = XSCRIPTCONTEXT .getComponentContext()
    except NameError:
        ctx = officehelper.bootstrap()
        if ctx is None:
            print("ERROR: Could not bootstrap default Office.")
            sys.exit(1)

# Starting from command line
if __name__ == "__main__":
    main()


# pythonloader loads a static g_ImplementationHelper variable
g_ImplementationHelper = unohelper.ImplementationHelper()
g_ImplementationHelper.addImplementation(
    Wavelet,  # UNO object class
    "org.extension.sample.do",  # implementation name (customize for yourself)
    ("com.sun.star.task.Job",), )  # implemented services (only 1)

    