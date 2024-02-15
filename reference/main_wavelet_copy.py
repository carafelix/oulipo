import uno
import unohelper
from com.sun.star.task import XJobExecutor
 
class Wavelet( unohelper.Base, XJobExecutor ):
    def __init__( self, ctx ):
        self.ctx = ctx
 
    def trigger( self, args ):
        desktop = self.ctx.ServiceManager.createInstanceWithContext(
            "com.sun.star.frame.Desktop", self.ctx )
 
        doc = desktop.getCurrentComponent()
 
        try:
            search = doc.createSearchDescriptor()
            search.SearchRegularExpression = True
            search.SearchString = "\\<(hola|pichi|v|z|o|u|i|a) "
 
            found = doc.findFirst( search )
            while found:
                found.String = found.String.replace( "hola", "sexobrutopapi" )
                found = doc.findNext( found.End, search)
 
        except:
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
