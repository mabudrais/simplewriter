/**
 *this code is distributed under BSD license
 * @author mohamed elimam abudrais 
 */
package com.example;

import com.sun.star.beans.PropertyValue;
import com.sun.star.beans.PropertyVetoException;
import com.sun.star.beans.UnknownPropertyException;
import com.sun.star.beans.XPropertySet;
import com.sun.star.comp.helper.BootstrapException;
import com.sun.star.frame.XDispatchHelper;
import com.sun.star.frame.XDispatchProvider;
import com.sun.star.frame.XFrame;
import com.sun.star.lang.IllegalArgumentException;
import com.sun.star.lang.WrappedTargetException;
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
import java.awt.Color;
import java.util.StringTokenizer;
import java.util.logging.Level;
import java.util.logging.Logger;

/*
 * To change this template, choose Tools | Templates
 * and open the template in the editor.
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
    private  XDispatchHelper xglobalDispatchHelper ;
    public final int write_aligment = 1;
    public final int left_aligment = 2;
    public final int center_aligment = 3;
    private XDispatchProvider xProvider;
    
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
        xText = xTextDocument.getText();
        xTCursor = xTextDocument.getText().createTextCursor();
        xParagraphCursor = (XParagraphCursor) UnoRuntime.queryInterface(XParagraphCursor.class, xTCursor);
          XFrame xFrame = xDesktop.getCurrentFrame();
            //Query for the frame's DispatchProvider
             xProvider = (XDispatchProvider) UnoRuntime.queryInterface(
                    XDispatchProvider.class, xFrame);
            XMultiServiceFactory xFactory = (XMultiServiceFactory) UnoRuntime.queryInterface(XMultiServiceFactory.class, xMCF);
            Object dispatchHelper = null;
        try {
            dispatchHelper = xFactory.createInstance("com.sun.star.frame.DispatchHelper");
        } catch (com.sun.star.uno.Exception ex) {
            Logger.getLogger(odf_document.class.getName()).log(Level.SEVERE, null, ex);
        }
            xglobalDispatchHelper = (XDispatchHelper) UnoRuntime.queryInterface(
                    XDispatchHelper.class, dispatchHelper);
    }

    public void insert_text(String txt) {
        try {
            xTextDocument.getText().insertString(xTCursor, txt, false);
           // xText.insertControlCharacter(xTCursor, com.sun.star.text.ControlCharacter.PARAGRAPH_BREAK, false);
           // set_font(txt);
        } catch (Exception ex) {
            Logger.getLogger(odf_document.class.getName()).log(Level.SEVERE, null, ex);
        }
    }
   public void select_text_to_left(int num_char_to_select){
       for(int i=0;i<num_char_to_select;i++){
        PropertyValue[] bPropertyValues = new PropertyValue[2];
            PropertyValue bPropertyValue1 = new PropertyValue();
            PropertyValue bPropertyValue2 = new PropertyValue();
            bPropertyValue1.Name="Count";
            bPropertyValue1.Value=1;
            bPropertyValue2.Name= "Select";
            bPropertyValue2.Value=true;
            bPropertyValues[0] = bPropertyValue1;
            bPropertyValues[1] = bPropertyValue2;
            xglobalDispatchHelper.executeDispatch(xProvider,  ".uno:GoLeft", "", 0, bPropertyValues);
       }
   }
    public void select_text_to_write(int num_char_to_select){
       for(int i=0;i<num_char_to_select;i++){
        PropertyValue[] bPropertyValues = new PropertyValue[2];
            PropertyValue bPropertyValue1 = new PropertyValue();
            PropertyValue bPropertyValue2 = new PropertyValue();
            bPropertyValue1.Name="Count";
            bPropertyValue1.Value=1;
            bPropertyValue2.Name= "Select";
            bPropertyValue2.Value=true;
            bPropertyValues[0] = bPropertyValue1;
            bPropertyValues[1] = bPropertyValue2;
            xglobalDispatchHelper.executeDispatch(xProvider,".uno:GoRight", "", 0, bPropertyValues);
       }
   }
    public void insert_paragraph(int aligment) {
        try {
            xText.insertControlCharacter(xTCursor, com.sun.star.text.ControlCharacter.PARAGRAPH_BREAK, false);
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
            set_aligment(left_aligment);
            XTextTable xTable = insert_emty_tabel(table);
            for (int x = 0; x < table.getCol_num(); x++) {
                for (int y = 0; y < table.getRow_num(); y++) {
                    insertIntoCell(table.get_cell_name(y, x), table.get_cell_text(y, x), xTable, 0);
                }
            }
        } catch (Exception e) {
            System.out.print(e.getLocalizedMessage());
        }
    }
   public void insert_write_left_table(odf_table table) {
        try {
            XTextTable xTable = insert_emty_tabel(table);
            for (int x = 0; x < table.getCol_num(); x++) {
                for (int y = 0; y < table.getRow_num(); y++) {
                    int xrihgt=table.getCol_num()-x-1;
                    if(table.get_cell_formula(y, x).length()==0) {
                        insertIntoCell(table.get_cell_name(y, xrihgt), table.get_cell_text(y, x), xTable, 0);
                    }else{
                        String str=table.get_cell_formula(y,x);
                        xTable.getCellByName(table.get_cell_name(y, xrihgt)).setFormula(str);
                    }
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
 public void set_char_colour(Color c) {
           PropertyValue[] bPropertyValues = new PropertyValue[1];
            PropertyValue bPropertyValue1 = new PropertyValue();
            bPropertyValue1.Name="FontColor";
            bPropertyValue1.Value=c.getRGB();
            bPropertyValues[0] = bPropertyValue1;
            xglobalDispatchHelper.executeDispatch(xProvider,  ".uno:FontColor", "", 0, bPropertyValues);
        }
 public void set_para_font(String font_name) {
     XPropertySet xparaProps = (XPropertySet) UnoRuntime.queryInterface(XPropertySet.class, xParagraphCursor);
        try {
            xparaProps.setPropertyValue("CharFontName",font_name);
        } catch (UnknownPropertyException ex) {
            Logger.getLogger(odf_document.class.getName()).log(Level.SEVERE, null, ex);
        } catch (PropertyVetoException ex) {
            Logger.getLogger(odf_document.class.getName()).log(Level.SEVERE, null, ex);
        } catch (IllegalArgumentException ex) {
            Logger.getLogger(odf_document.class.getName()).log(Level.SEVERE, null, ex);
        } catch (WrappedTargetException ex) {
            Logger.getLogger(odf_document.class.getName()).log(Level.SEVERE, null, ex);
        }
 }
    public void set_font(String font_name) {
        try {
             PropertyValue[] bPropertyValues = new PropertyValue[5];
            PropertyValue bPropertyValue1 = new PropertyValue();
            PropertyValue bPropertyValue2 = new PropertyValue();
            PropertyValue bPropertyValue3 = new PropertyValue();
            PropertyValue bPropertyValue4 = new PropertyValue();
            PropertyValue bPropertyValue5 = new PropertyValue();
            bPropertyValues[0] = bPropertyValue1;
            bPropertyValues[1] = bPropertyValue2;
            bPropertyValues[2] = bPropertyValue3;
            bPropertyValues[3] = bPropertyValue4;
            bPropertyValues[4] = bPropertyValue5;
            bPropertyValue1.Name = "CharFontName.StyleName";
            bPropertyValue1.Value = "";
            bPropertyValue2.Name = "CharFontName.Pitch";
            bPropertyValue2.Value = 2;
            bPropertyValue3.Name =  "CharFontName.CharSet";
            bPropertyValue3.Value = -1;
            bPropertyValue4.Name =  "CharFontName.Family";
            bPropertyValue4.Value = 5;
            bPropertyValue5.Name = "CharFontName.FamilyName";
            bPropertyValue5.Value = font_name;
            xglobalDispatchHelper.executeDispatch(xProvider,  ".uno:CharFontName", "", 0, bPropertyValues);
        } catch (Exception ex) {
        }
    }

    public void set_char_hight(int hight) {
        try {
            PropertyValue[] bPropertyValues = new PropertyValue[3];
            PropertyValue bPropertyValue1 = new PropertyValue();
            PropertyValue bPropertyValue2 = new PropertyValue();
            PropertyValue bPropertyValue3 = new PropertyValue();
            bPropertyValue1.Name = "FontHeight.Height";
            bPropertyValue1.Value = hight;
            bPropertyValue2.Name = "FontHeight.Prop";
            bPropertyValue2.Value = 100;
            bPropertyValue3.Name = "FontHeight.Diff";
            bPropertyValue3.Value = hight;
            bPropertyValues[0] = bPropertyValue1;
            bPropertyValues[1] = bPropertyValue2;
            bPropertyValues[2] = bPropertyValue3;
            xglobalDispatchHelper.executeDispatch(xProvider, ".uno:FontHeight", "", 0, bPropertyValues);
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
            //Thread.sleep(50);
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
          /* com.sun.star.table.ShadowFormat[] a= new com.sun.star.table.ShadowFormat[1];
            a[0].Location=com.sun.star.table.ShadowLocation.BOTTOM_LEFT;
            a[0].IsTransparent=false;
            a[0].ShadowWidth=500;
            a[0].Color=new Integer(8421504);
            xPropSet.setPropertyValue("ParaShadowFormat",a[0]);*/
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

    private XTextTable insert_emty_tabel(odf_table table) throws com.sun.star.uno.Exception, IllegalArgumentException {
        // Create a new table from the document's factory
        XTextTable xTable = (XTextTable) UnoRuntime.queryInterface(
                XTextTable.class, mxDocFactory.createInstance(
                "com.sun.star.text.TextTable"));
        int r = table.getRow_num();
        int c = table.getCol_num();
        xTable.initialize(r, c);
        xText.insertControlCharacter(xTCursor, com.sun.star.text.ControlCharacter.PARAGRAPH_BREAK, false);
        xText.insertTextContent(xTCursor, xTable, false);
        return xTable;
    }
}
