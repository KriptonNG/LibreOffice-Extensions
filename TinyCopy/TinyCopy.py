# This Source Code Form is subject to the terms of the Mozilla Public
# License, v. 2.0. If a copy of the MPL was not distributed with this
# file, You can obtain one at http://mozilla.org/MPL/2.0/.

import datetime

import uno, unohelper
from com.sun.star.task import XJobExecutor
from com.sun.star.document import XEventListener


class TinyCopy(unohelper.Base, XJobExecutor, XEventListener):


    def trigger(self, args):
        frame = self.desktop.ActiveFrame

        struct = uno.createUnoStruct('com.sun.star.beans.PropertyValue')
        struct.Name = 'ToPoint'
        struct.Value = 'Sheet1.A1'

        struct2 = uno.createUnoStruct('com.sun.star.beans.PropertyValue')
        struct2.Name = 'ToPoint'
        struct2.Value = 'Sheet1.B1'

        self.dispatchhelper.executeDispatch(
            frame,
            ".uno:GoToCell",
            "",
            0,
            tuple([struct]))

        self.dispatchhelper.executeDispatch(
            frame,
            ".uno:Copy",
            "",
            0,
            ())


        self.dispatchhelper.executeDispatch(
            frame,
            ".uno:GoToCell",
            "",
            0,
            tuple([struct2]))

        self.dispatchhelper.executeDispatch(
            frame,
            ".uno:Paste",
            "",
            0,
            ())


    def __init__(self, context):
        self.context = context
        
        self.desktop = self.createUnoService("com.sun.star.frame.Desktop")
        
        self.dispatchhelper = self.createUnoService("com.sun.star.frame.DispatchHelper")
    def createUnoService(self, name):
        
        return self.context.ServiceManager.createInstanceWithContext(name, self.context)
    def disposing(self, args):
        pass
    def notifyEvent(self, args):
        pass

g_ImplementationHelper = unohelper.ImplementationHelper()
g_ImplementationHelper.addImplementation(
    TinyCopy,
    "org.libreoffice.TinyCopy",
    ("com.sun.star.task.JobExecutor",))
