import uno
import unohelper

from com.sun.star.task import XJobExecutor

def iterate(xenum):
    if hasattr(xenum, 'hasMoreElements'):
            while xenum.hasMoreElements():
                yield xenum.nextElement()          

def xenumeration_list(xenum):
     return list(iterate(xenum))

class CommentJob(unohelper.Base, XJobExecutor):
    def __init__(self, ctx):
        self.ctx = ctx

    def trigger(self, args):
        comments = []
        desktop = self.ctx.ServiceManager.createInstanceWithContext(
            "com.sun.star.frame.Desktop", self.ctx)
        model = desktop.getCurrentComponent()


        URL = model.URL.rpartition('/')[0]
        text = model.getText()

        paragraphs = text.createEnumeration()
        paragraphs = xenumeration_list(paragraphs)

        for paragraph in paragraphs:
            services = paragraph.SupportedServiceNames
            if 'com.sun.star.text.Paragraph' in services:
                portions = xenumeration_list(paragraph.createEnumeration())
                for portion in portions:
                    #http://api.libreoffice.org/docs/idl/ref/servicecom_1_1sun_1_1star_1_1text_1_1textfield_1_1Annotation.html
                    if portion.TextPortionType == 'Annotation':
                            comments.append('%s \nAuthor: %s \n' % (portion.TextField.Content,portion.TextField.Author))

        outdoc = model.CurrentController.Frame.loadComponentFromURL("private:factory/swriter", "_blank", 0, ())
        outdoc_text = outdoc.Text
        outdoc_cursor = outdoc_text.createTextCursor()
        for comment in comments:
            outdoc_text.insertString(outdoc_cursor, comment+"\n", 0)
        outdoc.storeAsURL(URL + "/comments.doc",())
        outdoc.close(True)


        #Add Comment & Text
        #text = model.Text
        #cursor = text.createTextCursor()
        #comment = model.createInstance('com.sun.star.text.textfield.Annotation')
        ##http://api.libreoffice.org/docs/idl/ref/structcom_1_1sun_1_1star_1_1util_1_1Date.html
        #date = model.createInstance('com.sun.star.util.Date')
        #date.Day = 1
        #date.Month = 10
        #date.Year = 2016
        #comment.Content = 'test'
        #comment.Author = 'feyza'
        #comment.Date = date
        #text.insertTextContent(cursor,comment,False)

g_ImplementationHelper = unohelper.ImplementationHelper()

g_ImplementationHelper.addImplementation( \
    CommentJob,
    "org.libreoffice.Comment",
    ("com.sun.star.task.Job",),)
