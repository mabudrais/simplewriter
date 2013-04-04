package com.example;

import com.sun.star.beans.PropertyValue;
import com.sun.star.beans.XPropertySet;
import com.sun.star.comp.helper.BootstrapException;
import com.sun.star.frame.XDispatchHelper;
import com.sun.star.frame.XDispatchProvider;
import com.sun.star.frame.XFrame;
import com.sun.star.lang.IllegalArgumentException;
import com.sun.star.lang.XMultiServiceFactory;
import com.sun.star.style.ParagraphAdjust;
import com.sun.star.text.HoriOrientation;
import com.sun.star.text.VertOrientation;
import com.sun.star.text.WrapTextMode;
import com.sun.star.text.XParagraphCursor;
import com.sun.star.text.XText;
import com.sun.star.text.XTextCursor;
import com.sun.star.text.XTextRange;
import com.sun.star.text.XTextTable;
import com.sun.star.uno.UnoRuntime;
import java.util.logging.Level;
import java.util.logging.Logger;

/*
 * To change this template, choose Tools | Templates
 * and open the template in the editor.
 */
/**
 *this code is distributed under BSD license
 * @author „Õ„œ «·«„«„ «»Ê÷—Ì” 
 */

public class odf_document {

    private com.sun.star.text.XTextDocument xTextDocument = null;
    com.sun.star.frame.XDesktop xDesktop = null;
    com.sun.star.lang.XMultiComponentFactory xMCF = null;
    private XTextRange mxDocCursor;
    private XMultiServiceFactory mxDocFactory;
    private XParagraphCursor xParagraphCursor = null;
    private com.sun.star.text.XText xText = null;
    private com.sun.star.text.XTextCursor xTCursor = null;
    public final int write_aligment = 1;
    public final int left_aligment = 2;
    public final int center_aligment = 3;

    public odf_document() {

        xDesktop = getDesktop();

        try {
            com.sun.star.uno.XComponentContext xContext;

            // get the remote office component context
            xContext = com.sun.star.comp.helper.Bootstrap.bootstrap();

            // get the remote office service manager
            xMCF = xContext.getServiceManager();
            if (xMCF != null) {
                System.out.println("Connected to a running office ...");

                Object oDesktop = xMCF.createInstanceWithContext(
                        "com.sun.star.frame.Desktop", xContext);
                xDesktop = (com.sun.star.frame.XDesktop) UnoRuntime.queryInterface(
                        com.sun.star.frame.XDesktop.class, oDesktop);
            } else {
                System.out.println("Can't create a desktop. No connection, no remote office servicemanager available!");
            }
        } catch (BootstrapException | com.sun.star.uno.Exception e) {
            e.printStackTrace(System.err);
            System.exit(1);
        }
        xTextDocument = createTextdocument(xDesktop);
        mxDocFactory = (XMultiServiceFactory) UnoRuntime.queryInterface(
                XMultiServiceFactory.class, xTextDocument);
        xParagraphCursor = (XParagraphCursor) UnoRuntime.queryInterface(XParagraphCursor.class, xTCursor);
        xText = xTextDocument.getText();
        xTCursor = xTextDocument.getText().createTextCursor();
    }

    public void insert_text(String txt) {
        try {
            xTextDocument.getText().insertString(xTCursor, txt, false);
            xText.insertControlCharacter(xTCursor, com.sun.star.text.ControlCharacter.PARAGRAPH_BREAK, false);
            set_font(txt);
        } catch (IllegalArgumentException ex) {
            Logger.getLogger(odf_document.class.getName()).log(Level.SEVERE, null, ex);
        }
    }

    public void insert_paragraph(int aligment) {
        try {
            xParagraphCursor.gotoNextParagraph(true);
        } catch (Exception ex) {
            Logger.getLogger(odf_document.class.getName()).log(Level.SEVERE, null, ex);
        }
    }

    /**
     *
     * @param table
     */
    public void insert_table(odf_table table) {
        try {
            // Create a new table from the document's factory
            XTextTable xTable = (XTextTable) UnoRuntime.queryInterface(
                    XTextTable.class, mxDocFactory.createInstance(
                    "com.sun.star.text.TextTable"));
            int r = table.getRow_num();
            int c = table.getCol_num();
            xTable.initialize(r, c);
            xText.insertControlCharacter(xTCursor, com.sun.star.text.ControlCharacter.PARAGRAPH_BREAK, false);
            xText.insertTextContent(xTCursor, xTable, false);
            for (int x = 0; x < table.getCol_num(); x++) {
                for (int y = 0; y < table.getRow_num(); y++) {
                    insertIntoCell(table.get_cell_name(y, x), table.get_cell_text(y, x), xTable, 0);
                }
            }
        } catch (Exception e) {
            System.out.print(e.getLocalizedMessage());
        }
    }

    protected static void insertIntoCell(String sCellName, String sText,
            XTextTable xTable, int colour) {
        // Access the XText interface of the cell referred to by sCellName
        XText xCellText = (XText) UnoRuntime.queryInterface(
                XText.class, xTable.getCellByName(sCellName));

        // create a text cursor from the cells XText interface
        XTextCursor xCellCursor = xCellText.createTextCursor();
        // Get the property set of the cell's TextCursor
        XPropertySet xCellCursorProps = (XPropertySet) UnoRuntime.queryInterface(
                XPropertySet.class, xCellCursor);
        try {
            // Set the colour of the text to white
            xCellCursorProps.setPropertyValue("CharColor", new Integer(colour));
        } catch (Exception e) {
            e.printStackTrace();
        }
        // Set the text in the cell to sText
        xCellText.setString(sText);
    }

    public void set_font(String font_name) {
        try {
            XPropertySet xCursorProps = (XPropertySet) UnoRuntime.queryInterface(XPropertySet.class, xTCursor);
            xCursorProps.setPropertyValue("CharFontName", "Calibri");
            xCursorProps.setPropertyValue("ParaAdjust", ParagraphAdjust.CENTER);
            System.out.println(xCursorProps.getPropertyValue("CharFontName"));
        } catch (Exception ex) {
        }
    }

    public void set_char_hight(int hight) {
        try {
            XPropertySet xCursorProps = (XPropertySet) UnoRuntime.queryInterface(XPropertySet.class, xTCursor);
            xCursorProps.setPropertyValue("CharHeight", new Integer(hight));
        } catch (Exception ex) {
        }
    }
    public void set_aligment(int align) {
        try {
            XPropertySet xCursorProps = (XPropertySet) UnoRuntime.queryInterface(XPropertySet.class, xTCursor);
            switch (align) {
                case write_aligment:
                    xCursorProps.setPropertyValue("ParaAdjust", ParagraphAdjust.RIGHT);
                    break;
                case left_aligment:
                    xCursorProps.setPropertyValue("ParaAdjust", ParagraphAdjust.LEFT);
                    break;
                case center_aligment:
                    xCursorProps.setPropertyValue("ParaAdjust", ParagraphAdjust.CENTER);
                    break;
            }
        } catch (Exception ex) {
        }
    }

    public void insert_image(int x, int y,int w,int h,String file_name) {
        try {
            Object oGraphic = null;
            oGraphic = mxDocFactory.createInstance("com.sun.star.text.TextGraphicObject");
            com.sun.star.text.XTextContent xTextContent =
                    (com.sun.star.text.XTextContent) UnoRuntime.queryInterface(
                    com.sun.star.text.XTextContent.class, oGraphic);
            xText.insertTextContent(xTCursor, xTextContent, true);
            xTCursor.gotoEnd(true);
            com.sun.star.beans.XPropertySet xPropSet =
                    (com.sun.star.beans.XPropertySet) UnoRuntime.queryInterface(
                    com.sun.star.beans.XPropertySet.class, oGraphic);
            java.io.File sourceFile = new java.io.File(file_name);
            StringBuffer sUrl = new StringBuffer("file:///");
            sUrl.append(sourceFile.getCanonicalPath().replace('\\', '/'));
            System.out.println("insert graphic \"" + sUrl + "\"");
            // Setting the anchor type
            xPropSet.setPropertyValue("AnchorType",
                    com.sun.star.text.TextContentAnchorType.AT_PARAGRAPH);
             xPropSet.setPropertyValue("ContourOutside",false);
            // Setting the graphic url
            xPropSet.setPropertyValue("GraphicURL", sUrl.toString());
            WrapTextMode oWTM = null; // Choose among WrapTextMode elements 
            xPropSet.setPropertyValue("TextWrap", oWTM);
            // xPropSet.setPropertyValue("HoriOrient", HoriOrientation.NONE); 
            //xPropSet.setPropertyValue("VertOrient", VertOrientation.NONE);
            //xPropSet.setPropertyValue( "HoriOrientPosition",  new Integer( 15500 ) );
           // xPropSet.setPropertyValue("VertOrientPosition", new Integer(4200));
            if(w>0) {
                xPropSet.setPropertyValue("Width", new Integer(w));
            }
            if(h>0) {
                xPropSet.setPropertyValue("Height", new Integer(h));
            }
            if(x>-1){
                xPropSet.setPropertyValue("HoriOrient", HoriOrientation.NONE); 
                xPropSet.setPropertyValue( "HoriOrientPosition",  new Integer(x) );
            }
            if(y>-1){
                xPropSet.setPropertyValue("VertOrient", VertOrientation.NONE);
                xPropSet.setPropertyValue("VertOrientPosition", new Integer(y));
            }
        } catch (Exception ex) {
        }
    }
    public void set_paragraph_aligment(int aligment) {
        try {
            XFrame xFrame = xDesktop.getCurrentFrame();
            //Query for the frame's DispatchProvider
            XDispatchProvider xProvider = (XDispatchProvider) UnoRuntime.queryInterface(
                    XDispatchProvider.class, xFrame);
            XMultiServiceFactory xFactory = (XMultiServiceFactory) UnoRuntime.queryInterface(XMultiServiceFactory.class, xMCF);
            Object dispatchHelper = xFactory.createInstance("com.sun.star.frame.DispatchHelper");
            XDispatchHelper xDispatchHelper = (XDispatchHelper) UnoRuntime.queryInterface(
                    XDispatchHelper.class, dispatchHelper);
            PropertyValue[] bPropertyValues = new PropertyValue[1];
            PropertyValue bPropertyValue = new PropertyValue();
            bPropertyValue.Value = true;
            bPropertyValues[0] = bPropertyValue;
            switch (aligment) {
                case 1:
                    bPropertyValue.Name = "RightPara";
                    xDispatchHelper.executeDispatch(xProvider, ".uno:RightPara", "", 0, bPropertyValues);
                    break;
            }
        } catch (com.sun.star.uno.Exception ex) {
            Logger.getLogger(odf_document.class.getName()).log(Level.SEVERE, null, ex);
        }
    }

    public static com.sun.star.text.XTextDocument createTextdocument(
            com.sun.star.frame.XDesktop xDesktop) {
        com.sun.star.text.XTextDocument aTextDocument = null;

        try {
            com.sun.star.lang.XComponent xComponent = CreateNewDocument(xDesktop,
                    "swriter");
            aTextDocument = (com.sun.star.text.XTextDocument) UnoRuntime.queryInterface(
                    com.sun.star.text.XTextDocument.class, xComponent);
        } catch (Exception e) {
            e.printStackTrace(System.err);
        }

        return aTextDocument;
    }

    protected static com.sun.star.lang.XComponent CreateNewDocument(
            com.sun.star.frame.XDesktop xDesktop,
            String sDocumentType) {
        String sURL = "private:factory/" + sDocumentType;

        com.sun.star.lang.XComponent xComponent = null;
        com.sun.star.frame.XComponentLoader xComponentLoader = null;
        com.sun.star.beans.PropertyValue xValues[] =
                new com.sun.star.beans.PropertyValue[1];
        com.sun.star.beans.PropertyValue xEmptyArgs[] =
                new com.sun.star.beans.PropertyValue[0];

        try {
            xComponentLoader = (com.sun.star.frame.XComponentLoader) UnoRuntime.queryInterface(
                    com.sun.star.frame.XComponentLoader.class, xDesktop);

            xComponent = xComponentLoader.loadComponentFromURL(
                    sURL, "_blank", 0, xEmptyArgs);
        } catch (Exception e) {
            e.printStackTrace(System.err);
        }

        return xComponent;
    }

    public static com.sun.star.frame.XDesktop getDesktop() {
        com.sun.star.frame.XDesktop xDesktop = null;
        com.sun.star.lang.XMultiComponentFactory xMCF = null;

        try {
            com.sun.star.uno.XComponentContext xContext = null;

            // get the remote office component context
            xContext = com.sun.star.comp.helper.Bootstrap.bootstrap();

            // get the remote office service manager
            xMCF = xContext.getServiceManager();
            if (xMCF != null) {
                System.out.println("Connected to a running office ...");

                Object oDesktop = xMCF.createInstanceWithContext(
                        "com.sun.star.frame.Desktop", xContext);
                xDesktop = (com.sun.star.frame.XDesktop) UnoRuntime.queryInterface(
                        com.sun.star.frame.XDesktop.class, oDesktop);
            } else {
                System.out.println("Can't create a desktop. No connection, no remote office servicemanager available!");
            }
        } catch (Exception e) {
            e.printStackTrace(System.err);
            System.exit(1);
        }


        return xDesktop;
    }
}
