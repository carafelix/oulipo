import uno
import unohelper
from com.sun.star.task import XJobExecutor
from com.sun.star.awt import XKeyHandler
from com.sun.star.awt import Key
from com.sun.star.text import XTextDocument
from com.sun.star.lang import XEventListener
from com.sun.star.awt.MessageBoxButtons import BUTTONS_OK
from com.sun.star.awt.MessageBoxType import INFOBOX

 
class Wavelet( unohelper.Base, XJobExecutor, XEventListener ):
    def __init__( self, ctx ):
        self.ctx = ctx

        desktop = ctx.getServiceManager().createInstanceWithContext("com.sun.star.frame.Desktop", ctx)
        self.doc = desktop.getCurrentComponent()

        self.doc.Text.addEventListener(wavelet)


    def documentEventOccured(self, event):
        if event.EventName == "InsertText":
            self.trigger(None)

    def get_cursor_on_last_modified_word(self):
        if self.doc is None:
            return
        
        text = self.doc.Text

        # Get the view cursor
        view_cursor = self.doc.CurrentController.getViewCursor()
        
        # Get the position of the last modified character

        last_modified_position = view_cursor.getStart()

        # Get the word at the last modified position

        cursor = text.createTextCursorByRange(last_modified_position)
        cursor.gotoStartOfWord(False)
        cursor.gotoEndOfWord(True)
        last_modified_word = cursor.getString()
        
        return cursor
    
    def trigger( self, args ):
        if self.doc is None:
            return

        last_word_cursor = self.get_cursor_on_last_modified_word(doc)

        try:
            search = self.doc.createSearchDescriptor()
            search.SearchWords = True
            search.SearchString = last_word_cursor.getString()

            found = self.doc.findAll( search )
            
            if(found.Count > 1):
                last_word_cursor.CharBackColor = 0xFFFF00
                # for i in range(found.Count):
                #     found.getByIndex(i).CharBackColor = 0xFFFF00
            else:
                # since we are only selecting tangent words, this doesn't get updated when you erase the duplicate word
                last_word_cursor.CharBackColor = -1
    
        except Error as e:
            show_message_box(e, "Error")
            pass

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


# pythonloader loads a static g_ImplementationHelper variable
g_ImplementationHelper = unohelper.ImplementationHelper()
g_ImplementationHelper.addImplementation(
    Wavelet,  # UNO object class
    "org.extension.sample.do",  # implementation name (customize for yourself)
    ("com.sun.star.task.Job",), )  # implemented services (only 1)

    




# def MessageBox(ParentWindow, MsgText, MsgTitle, MsgType, MsgButtons):
#     ctx = XSCRIPTCONTEXT.getComponentContext()
#     sm = ctx.ServiceManager
#     si = sm.createInstanceWithContext("com.sun.star.awt.Toolkit", ctx)
#     mBox = si.createMessageBox(ParentWindow, MsgType, MsgButtons, MsgTitle, MsgText)
#     mBox.execute()