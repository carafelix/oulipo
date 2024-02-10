# -*- tab-width: 4; indent-tabs-mode: nil; py-indent-offset: 4 -*-
#
# This file is part of the LibreOffice project.
#
# This Source Code Form is subject to the terms of the Mozilla Public
# License, v. 2.0. If a copy of the MPL was not distributed with this
# file, You can obtain one at http://mozilla.org/MPL/2.0/.
#

import sys
import unohelper
import officehelper

from com.sun.star.task import XJobExecutor


# The MainJob is a UNO component derived from unohelper.Base class
# and also the XJobExecutor, the implemented interface
class MainJob(unohelper.Base, XJobExecutor):
    def __init__(self, context):
        self.context = context
        # handling different situations (inside LibreOffice or other process)
        try:
            self.sm = context.getServiceManager()
            self.desktop = XSCRIPTCONTEXT.getDesktop()
        except NameError:
            self.sm = context.ServiceManager
            self.desktop = self.context.getServiceManager().createInstanceWithContext(
                "com.sun.star.frame.Desktop", self.context)

    def trigger(self, args):
        desktop = self.context.ServiceManager.createInstanceWithContext(
            "com.sun.star.frame.Desktop", self.context)
        model = desktop.getCurrentComponent()
        if not hasattr(model, "Text"):
            model = self.desktop.loadComponentFromURL("private:factory/swriter", "_blank", 0, ())
        text = model.Text
        cursor = text.createTextCursor()
        text.insertString(cursor, "Hello fukin fokers -> " + args + "\n", 0)


# Starting from Python IDE
def main():
    try:
        context = XSCRIPTCONTEXT
    except NameError:
        context = officehelper.bootstrap()
        if context is None:
            print("ERROR: Could not bootstrap default Office.")
            sys.exit(1)
    job = MainJob(context)
    job.trigger()


# Starting from command line
if __name__ == "__main__":
    main()


# pythonloader loads a static g_ImplementationHelper variable
g_ImplementationHelper = unohelper.ImplementationHelper()

g_ImplementationHelper.addImplementation(
    MainJob,  # UNO object class
    "org.extension.sample.do",  # implementation name (customize for yourself)
    ("com.sun.star.task.Job",), )  # implemented services (only 1)

# vim: set shiftwidth=4 softtabstop=4 expandtab:
