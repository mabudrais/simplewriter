simple writer is a java classes for creating an open office writer documents.
it depends on oosdk.
to use the lib (i assume you are using netbeans):
1- install openoffice or (libreoffice)
2-download and install oosdk.
3-run the setsdkenv_windows.bat.
4-download and install OpenOffice.org API plugin for NetBeans.
5-add the file odf_document.java and odf_table.java to your project.
6-add juh.jar and ridl.jar to your project.

now when you need to creat open office writer doc:
odf_document doc = new odf_document(); //creating a new doc
doc.insert_text("Hello World"); //insrting some text to the document


