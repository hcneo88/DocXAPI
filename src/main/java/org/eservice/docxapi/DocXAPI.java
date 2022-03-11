package org.eservice.docxapi ;

//Reference : https://github.com/plutext/DocXAPI/tree/DocXAPI-3.2.2/src/samples/DocXAPI/org/DocXAPI/samples

//As at 17 Feb 2023 - STILL NOT WORKING
//1. Page numbering in Page x of y format.
//2. Insert paragraph - unable to figure out how to get the index.
//3. Not tested yet : setMargin and setDocumentMargin

//Defunct as at 28 Feb 2022
//1. createStdHeader  working but only left or right justified <= use createHeaderObject
//2. createStdFooter working but only left or right justified <= use createFooterObject

import org.apache.pdfbox.pdmodel.PDDocument;
import org.apache.pdfbox.pdmodel.encryption.AccessPermission;
import org.apache.pdfbox.pdmodel.encryption.StandardProtectionPolicy;
import org.docx4j.wml.HdrFtrRef;
import org.docx4j.Docx4J;
import org.docx4j.TraversalUtil;
import org.docx4j.XmlUtils;
import org.docx4j.finders.RangeFinder;
import org.docx4j.finders.SectPrFinder;
import org.docx4j.jaxb.Context;
import org.docx4j.model.datastorage.XPathEnhancerParser.qName_return;
import org.docx4j.model.fields.FieldUpdater;
import org.docx4j.model.structure.HeaderFooterPolicy;
import org.docx4j.model.structure.PageDimensions;
import org.docx4j.model.structure.SectionWrapper;
import org.docx4j.model.table.TblFactory;
import org.docx4j.dml.wordprocessingDrawing.Inline;
import org.docx4j.openpackaging.packages.ProtectDocument;
import org.docx4j.openpackaging.packages.WordprocessingMLPackage;
import org.docx4j.openpackaging.parts.Part;
import org.docx4j.openpackaging.parts.PartName;
import org.docx4j.openpackaging.parts.WordprocessingML.BinaryPart;
import org.docx4j.openpackaging.parts.WordprocessingML.BinaryPartAbstractImage;
import org.docx4j.openpackaging.parts.WordprocessingML.FooterPart;
import org.docx4j.openpackaging.parts.WordprocessingML.HeaderPart;
import org.docx4j.openpackaging.parts.WordprocessingML.ImageGifPart;
import org.docx4j.openpackaging.parts.WordprocessingML.ImageJpegPart;
import org.docx4j.openpackaging.parts.WordprocessingML.ImagePngPart;
import org.docx4j.openpackaging.parts.WordprocessingML.MainDocumentPart;
import org.docx4j.openpackaging.parts.relationships.Namespaces;
import org.docx4j.openpackaging.parts.relationships.RelationshipsPart;
import org.docx4j.openpackaging.parts.relationships.RelationshipsPart.AddPartBehaviour;
import org.docx4j.wml.*;
import org.docx4j.wml.PPrBase.Spacing;
import org.docx4j.wml.SectPr.PgMar;
import org.slf4j.Logger;
import org.slf4j.LoggerFactory;

import javax.xml.bind.JAXBElement;
import javax.xml.bind.JAXBException;

import com.google.zxing.BarcodeFormat;
import com.google.zxing.EncodeHintType;
import com.google.zxing.MultiFormatWriter;
import com.google.zxing.client.j2se.MatrixToImageWriter;
import com.google.zxing.common.BitMatrix;
import com.google.zxing.qrcode.decoder.ErrorCorrectionLevel;

import java.io.ByteArrayOutputStream;
import java.io.File;
import java.io.FileInputStream;
import java.io.FileOutputStream;
import java.io.FileWriter;
import java.io.IOException;
import java.io.InputStream;
import java.io.Writer;
import java.lang.reflect.Type;
import java.math.BigInteger;
import org.docx4j.org.apache.poi.util.IOUtils;
import org.docx4j.relationships.Relationship;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.EnumMap;
import java.util.HashMap;

import java.util.List;
import java.util.Map;
import java.util.Random;

import com.google.gson.Gson;
import com.google.gson.GsonBuilder;
import com.google.gson.internal.LinkedTreeMap;
import com.google.gson.reflect.TypeToken;

import lombok.extern.slf4j.Slf4j;

@Slf4j
public class DocXAPI { 
    
    class HeaderFooter {
        //NOTE: If this class is changed, need to analyze the impact on the method
        //deserializeHeaderFooter.  

        private String hdrFtrRef ;
        private List<List <Object>> paragraphs = new ArrayList<>() ;  // need to convert back to list<object>

        HeaderFooter() {}

        public void setParagraphs (List<List<Object>> pList) {
            paragraphs = pList;
        }

        public void setHdrFtrRef(HdrFtrRef hfRef) {
            if (hfRef == HdrFtrRef.FIRST) hdrFtrRef = "FIRST";
            if (hfRef == HdrFtrRef.EVEN) hdrFtrRef = "EVEN";
            if (hfRef == HdrFtrRef.DEFAULT) hdrFtrRef = "DEFAULT";

        }

        public List<List<Object>> getParagraphs() {
            return paragraphs;
        }

        public HdrFtrRef getHdrFtrRef() {
            HdrFtrRef hfRef = HdrFtrRef.DEFAULT;

            if (hdrFtrRef.equalsIgnoreCase("FIRST")) hfRef = HdrFtrRef.FIRST;
            if (hdrFtrRef.equalsIgnoreCase("EVEN")) hfRef = HdrFtrRef.EVEN;

            return hfRef ;
        }

    }
   
    class MergeFields {
                           //an array element is a "paragraph" in the header/footer 

        RecipientAddress recipientAddress  ;
        Map<String, TextField> textVariables ;              //1st String = name of merge field
        Map<String, ImageField> imageVariables  ;
        Map<String, List<Map<String, String>>> tables ;     //1st String = name of table
        //String keySeed = "" ;

        MergeFields() {

            textVariables = new HashMap<>();
            imageVariables = new HashMap<>();
            tables = new HashMap<>() ;

            recipientAddress = new RecipientAddress();
            recipientAddress.setRecipientName("");
            recipientAddress.setAddressType(' ');
            recipientAddress.setBlkHseNo("");
            recipientAddress.setBuildingName("");
            recipientAddress.setFloorNo("");
            recipientAddress.setStreetName("");
            recipientAddress.setPostalCd("");
            recipientAddress.setTransactionId("");
            recipientAddress.setTransactionDate("");
        } 
       
        public RecipientAddress gRecipientAddress() {
            return this.recipientAddress ;
        }

        public void setRecipientAddress(RecipientAddress recipientAddress) {
            this.recipientAddress = recipientAddress ;
        }
        
        public void setField(String fieldName, TextField fieldValue) {
//            textVariables.put(Constant.FLD_OPEN_DELIMITER + fieldName + Constant.FLD_CLOSE_DELIMITER , fieldValue) ;
            textVariables.put(fieldName , fieldValue) ;
        }

        public void setField(String fieldName, ImageField imageField) {
            String key = Constant.FLD_OPEN_DELIMITER + fieldName + Constant.FLD_CLOSE_DELIMITER ; 
            imageVariables.put(key, imageField) ;
        }

        public void setTableRowFields(String tableName, Map<String, String> rowFields ) {
                       
            List<Map<String, String>> tableList ;
            if (tables.containsKey(tableName)) {
                tableList = tables.get(tableName) ;               
            } else {
                tableList = new ArrayList<Map<String, String>>();
            }

            Map<String, String> mergeFieldMap = new HashMap<>() ; 
            for (Map.Entry<String,String> field : rowFields.entrySet()) {
                mergeFieldMap.put( Constant.FLD_OPEN_DELIMITER + field.getKey() + Constant.FLD_CLOSE_DELIMITER, field.getValue());
            }
            
            tableList.add(mergeFieldMap) ;            
            mergeField.tables.put(tableName, tableList) ;
        }
        
    }

    private static Logger log = LoggerFactory.getLogger(DocXAPI.class);
    
        
    private WordprocessingMLPackage templatePackage ; 
    private MainDocumentPart templateMainDocumentPart ;   
    private ObjectFactory objectFactory ;
   
    MergeFields mergeField = new MergeFields();
    Map<String, HeaderFooter> headerFooters = new HashMap<>();       
        
    public MainDocumentPart getTemplateMainDocumentPart() {
        return templateMainDocumentPart;
    }

    public void loadTemplate(String templateFileName) throws Exception {
        
        objectFactory = Context.getWmlObjectFactory();
        templatePackage = WordprocessingMLPackage.load(new File(templateFileName));
        templateMainDocumentPart = templatePackage.getMainDocumentPart();
        changeGlobalFont("Arial");        
        log.debug ("Loaded Template: {} {}", templateFileName,  " loaded.");
    } 
   
    public byte[] loadImage(String imagePath) throws Exception{
        InputStream is = new FileInputStream(imagePath);
        byte[] bytes = IOUtils.toByteArray(is);
        return bytes;
    }

    public Boolean nothingToMerge() {

        if (mergeField == null) return true ;
            
        return (mergeField.textVariables.size() == 0 && 
                mergeField.imageVariables.size() == 0 && 
                mergeField.tables.size() == 0 && 
                headerFooters.size()==0) ;
    }

    public void setHeaderFooter(String type, HdrFtrRef hdrFtrRef,       //type: "HEADER", "FOOTER"
                                List<List<Object>> paragraphs) {
        
        
        List<List<Object>> paraList = new ArrayList<List<Object>>() ;
        for (List<Object> objList : paragraphs) {

            List<Object> paraLine = new ArrayList<>() ;
            for (Object obj : objList) {
                if (obj instanceof byte[]) {
                    ImageField imageField = createImageField((byte[])obj).setSize(0) ;
                    paraLine.add(imageField) ;
                } else {

                    if (obj instanceof java.lang.String) {
                        TextField textField = createTextField((String) obj) ;
                        paraLine.add(textField) ;
                    } else {
                        paraLine.add(obj) ;
                    }
                }

            } //For obj
            paraList.add(paraLine) ;

        } //For objList

        HeaderFooter hf = new HeaderFooter() ;
        hf.setHdrFtrRef(hdrFtrRef);
        hf.setParagraphs(paraList);

        String key = type + "-";
        if (hdrFtrRef == HdrFtrRef.DEFAULT) key = key + "DEFAULT" ;
        if (hdrFtrRef == HdrFtrRef.FIRST) key = key + "FIRST" ;
        if (hdrFtrRef == HdrFtrRef.EVEN) key = key + "EVEN" ;
        headerFooters.put(key, hf);
    }

    
    public int setTableMergeFields(String tableName, Map <String, String> fieldMap) {
                
        if (mergeField == null) {  
            mergeField = new MergeFields();
        } 
        
        //Todo: validation to make sure field is defined
        
        mergeField.setTableRowFields(tableName, fieldMap) ;
        return 0 ;
       
    }

    public int setImageMergeFields (Map <String, ImageField> fieldMap) {

        if (mergeField == null) {  
            mergeField = new MergeFields();
        }

         //Todo: validation to make sure field is defined
        
        for (Map.Entry<String, ImageField> field : fieldMap.entrySet()) {
            mergeField.setField(field.getKey(), field.getValue())  ;
        }
        return 0;

    }   

    public int setTextMergeFields(Map <String, TextField> fieldMap) throws Exception {
        
           
        if (mergeField == null) {  
            mergeField = new MergeFields();
        } 
        
        //Todo: validation to make sure field is defined

        for (Map.Entry<String,TextField> field : fieldMap.entrySet()) {
            mergeField.setField(field.getKey(), field.getValue())  ;
        }
        
        return 0 ;
       
    }

    public RecipientAddress getRecipient() {
        return mergeField.gRecipientAddress() ;
    }

    public void setRecipient(RecipientAddress recipientAddress) {
        mergeField.setRecipientAddress(recipientAddress); 

        Map<String, String>  address = recipientAddress.format() ;
        mergeField.setField("transactionid", createTextField(recipientAddress.getTransactionId())) ;
        mergeField.setField("transactiondate", createTextField(recipientAddress.getTransactionDate())) ;
        mergeField.setField("recipientname", createTextField(address.get("recipientname")));
        mergeField.setField("addressline1",  createTextField(address.get("addressline1"))) ;
        mergeField.setField("addressline2",  createTextField(address.get("addressline2"))) ;
        mergeField.setField("addressline3",  createTextField(address.get("addressline3"))) ;


    }

    public String serializeMergeFields ()  {
        Gson gson = new GsonBuilder().setPrettyPrinting().create();
        String mergeString = gson.toJson(mergeField) ;         
        return mergeString ; 
    } 

    public String serializeMergeFields (String filePath) throws Exception  {
        Gson gson = new GsonBuilder().setPrettyPrinting().create();
        Writer jsonWriter =  new FileWriter(filePath) ;
        gson.toJson(mergeField, jsonWriter) ;
        jsonWriter.flush();
        jsonWriter.close();
        return gson.toJson(mergeField);          
    } 

    public void deserializeMergeFields(String json) {
        
        Gson gson = new Gson();
        Type classObject = new TypeToken<MergeFields>() {}.getType();
        mergeField = gson.fromJson(json, classObject);//mergeFieldClassObject) ;
    } 

    public String serializeHeaderFooter () throws Exception  {
        Gson gson = new GsonBuilder().setPrettyPrinting().create();
        String jsonString = gson.toJson(headerFooters) ; 
        
        return jsonString ;  
              
    } 

    public String serializeHeaderFooter (String filePath) throws Exception  {
        Gson gson = new GsonBuilder().setPrettyPrinting().create();
        Writer jsonWriter =  new FileWriter(filePath) ;
        gson.toJson(headerFooters, jsonWriter) ; 
        jsonWriter.flush();
        jsonWriter.close();
        return gson.toJson(headerFooters) ;
    } 


    int classType ;
    public void deserializeHeaderFooters(String json) throws Exception {
        
        headerFooters.clear() ;

        Gson gson = new GsonBuilder().setPrettyPrinting().create();
        Type objType = new TypeToken<Map<String, HeaderFooter>>(){}.getType();        
        Map<String, HeaderFooter> hf =  gson.fromJson(json, objType );

        for (Map.Entry<String, HeaderFooter> m : hf.entrySet()) {
            List<List<Object>> paraList = m.getValue().getParagraphs() ;  //getValue().getParagraphs == List<Object>
            
            List<List<Object>> paragraphs = new ArrayList<List<Object>>() ;
            List<Object> paraLine = null ;
            
            for (List<Object> pList : paraList) {

                paraLine = new ArrayList<>() ;
                for (Object o : pList) {

                    if (o == null) {
                        paraLine.add(null) ;
                        continue;
                    }

                    if (o instanceof LinkedTreeMap) { 
                        LinkedTreeMap<String,Object> treeMap  = (LinkedTreeMap) o;
                        TextField txt = new TextField();            
                        ImageField img = new ImageField();
                        
                        treeMap.entrySet().stream().forEach(e -> {

                            classType = 0 ;
                            if (o.toString().contains("formats")) {    
                                
                                classType = 1 ;
                                if (e.getKey().equalsIgnoreCase("value")) {
                                    txt.setText((String) e.getValue()) ;
                                }
                                if (e.getKey().equalsIgnoreCase("formats")) {
                                    txt.setFormats((String) e.getValue()) ;
                                }
                            } //TextField

                            if (o.toString().contains("image")) {
                                classType = 2 ;
                                if (e.getKey().equalsIgnoreCase("image")) {
                                    img.setImage((String) e.getValue()) ; 
                                }

                                if (e.getKey().equalsIgnoreCase("properties")) {
                                    Map<String, Object>  propMap  = (Map<String, Object>) e.getValue();
                                    img.setProperties(propMap) ;
                                }
                            } //ImageField

                        }); 
                        
                        if (classType == 1) paraLine.add(txt) ;
                        if (classType == 2) paraLine.add(img) ;

                    } //Object is a linkedTreeMap 
                } //For o
                paragraphs.add(paraLine) ;            
            }
            
            HeaderFooter headFoot = new HeaderFooter();
            headFoot.setHdrFtrRef(m.getValue().getHdrFtrRef());
            headFoot.setParagraphs(paragraphs);
            headerFooters.put(m.getKey(), headFoot) ;
            //System.out.println("HEADERFOOTER:" + headerFooters) ;
        } //Map
        
    }

    public void breakPage() {    
        
        P p = objectFactory.createP();        
     //   R r = objectFactory.createR();
        R r = createRun();
        p.getContent().add(r);       
        
        Br br = objectFactory.createBr();
        r.getContent().add(br);
        br.setType(org.docx4j.wml.STBrType.PAGE);
        templateMainDocumentPart.addObject(p);
        
    }

    public void saveDoc(String fileName) throws Exception {

        //FieldUpdater updater = new FieldUpdater(templatePackage);
		//updater.update(true);
        templatePackage.save(new File(fileName)) ;

    } 

    
    public void saveAsPDF(String docFileName, String pdfFileName) throws Exception {
        
        WordprocessingMLPackage docFile = WordprocessingMLPackage.load(new File(docFileName));
        
        if (pdfFileName == null) 
            pdfFileName = docFileName.replaceAll(".docx", ".pdf") ;
        FileOutputStream pdfOutputStream = new FileOutputStream(pdfFileName);
        Docx4J.toPDF (docFile, pdfOutputStream);
        pdfOutputStream.flush() ;
        pdfOutputStream.close() ;
      
      
   /*   if (pdfFileName == null) 
        pdfFileName = docFileName.replaceAll(".docx", ".pdf") ;
    
        InputStream doc = new FileInputStream(new File(docFileName));
        XWPFDocument document = new XWPFDocument(doc);
        PdfOptions options = PdfOptions.create();
        FileOutputStream out = new FileOutputStream(new File(pdfFileName));
        PdfConverter.getInstance().convert(document, out, options); */
    }
    
    public ByteArrayOutputStream saveAsPDF(String docFileName) throws Exception {

        WordprocessingMLPackage docPackage = WordprocessingMLPackage.load(new File(docFileName));

        ByteArrayOutputStream bos = new ByteArrayOutputStream(); 
        Docx4J.toPDF (docPackage, bos);
        bos.flush() ;
        return bos;
    }
    
    public void protectDoc(String passWord) {
        ProtectDocument protection = new ProtectDocument(templatePackage);
        protection.restrictEditing(STDocProtect.READ_ONLY, passWord);
    } 

    public void protectPDF(Object pdfContent, String encryptedPdfFileName, 
                           String ownerPwd, String userPwd) throws IOException {

        PDDocument pdf = null ;
        if (pdfContent instanceof java.lang.String) {
            String pdfFilename = (String)pdfContent ;
                pdf = PDDocument.load(new File(pdfFilename));
        } else {
            if (pdfContent instanceof byte[]) {
                byte[] pdfByte = (byte[])pdfContent;
                pdf = PDDocument.load(pdfByte) ;
            }
        }

        if (pdf == null) return ;

        // Define the length of the encryption key.
        // Possible values are 40, 128 or 256.
        int keyLength = 256;

        AccessPermission ap = new AccessPermission();
        // disable printing, 
        ap.setCanPrint(false);
        //disable copying
        ap.setCanExtractContent(false);
        //Disable other things if needed...

        // Owner password (to open the file with all permissions) is "12345"
        // User password (to open the file but with restricted permissions, is empty here)
        StandardProtectionPolicy spp = new StandardProtectionPolicy(ownerPwd, userPwd, ap);
        spp.setEncryptionKeyLength(keyLength);

        //Apply protection
        pdf.protect(spp);

        pdf.save(encryptedPdfFileName);
        pdf.close();
    }

    public PPr setSingleLineSpacing(PPr pProperties) {
 
        Spacing lineSpacing = objectFactory.createPPrBaseSpacing();
        lineSpacing.setAfterLines( BigInteger.ONE );
        lineSpacing.setAfter(BigInteger.ONE );
        pProperties.setSpacing(lineSpacing);
        return pProperties ;

    }

    /*
    public void setParagraphAlignment(int option) {
        PARAGRAPH_ALIGNMENT_OPTION = option ;
    }

    public int getParagraphAlignment() {
        return PARAGRAPH_ALIGNMENT_OPTION ;
    }

    public void setParagraphTabPosition(double cmFromLeft) {
        RIGHT_TAB_POSITION = cmFromLeft ;
    } 

    public double getParagraphTabPosition() {
        return RIGHT_TAB_POSITION; 
    } */ 

    public void justifyParagraph(P para, JcEnumeration alignment) { 
        
        PPr pProperties = para.getPPr() ;
        if (pProperties == null) pProperties = objectFactory.createPPr() ;
        
        Jc justification = objectFactory.createJc();
        justification.setVal(alignment);
        pProperties.setJc(justification);

        para.setPPr(setSingleLineSpacing(pProperties));
      //  return para ;

    }

    //https://www.DocXAPIava.org/forums/docx-java-f6/tabstops-t1705.html
    public void defineTabStop(P paragraph, double cm) {  // 1 cm = 574 unit
      
        PPr ppr = paragraph.getPPr() ; // objectFactory.createPPr();
        if (ppr == null) ppr = objectFactory.createPPr() ;
 
        CTTabStop tabstop = objectFactory.createCTTabStop();
        tabstop.setVal(org.docx4j.wml.STTabJc.LEFT);
        tabstop.setPos( BigInteger.valueOf(Math.round(cm*574))) ;

        Tabs tabs = objectFactory.createTabs();
        tabs.getTab().add( tabstop);
        ppr.setTabs(tabs);
        paragraph.setPPr(ppr);

    }

    public PgMar setMargins(long top, long right, long left, long bottom){
        PgMar pgMar = new PgMar();
        pgMar.setTop( BigInteger.valueOf(top));
        pgMar.setBottom( BigInteger.valueOf(bottom));
        pgMar.setLeft( BigInteger.valueOf(left));
        pgMar.setRight( BigInteger.valueOf(right));
        return pgMar ;   }
    
    public void setDocumentMargin() {
        //https://www.DocXAPIava.org/forums/docx-java-f6/how-to-set-page-margins-t1163.html

        MainDocumentPart mainDocumentPart = templatePackage.getMainDocumentPart();   
        Body body = mainDocumentPart.getJaxbElement().getBody();
        PageDimensions page = new PageDimensions();
        PgMar pgMar = page.getPgMar();                       
        SectPr sectPr = objectFactory.createSectPr();   
        body.setSectPr(sectPr);          
        sectPr.setPgMar(pgMar);   
    }

    public R rTabTo(R run, STPTabAlignment alignment) {
 
        R.Ptab tab = new R.Ptab();
        tab.setAlignment(alignment);
        tab.setRelativeTo(STPTabRelativeTo.MARGIN);
        tab.setLeader(STPTabLeader.NONE);
        if (run == null) run = createRun();
        run.getContent().add(tab)  ;
        return run;
             
        //return para ;
    }

    public R spaceTo(R run, int spaceCount) {
 
       // if (run == null) run = objectFactory.createR() ;
        if (run == null) run = createRun() ;
        Text text = objectFactory.createText();
        text.setSpace("preserve");
        
        StringBuilder sbSpace = new StringBuilder() ;
        for (int i=0;i< spaceCount;i++) {
            sbSpace.append(' ') ;
        } 
        text.setValue(sbSpace.toString());
        run.getContent().add(text) ;
               
        return run ;
             
        //return para ;
    }

    public String getXML(Object element) {
        return XmlUtils.marshaltoString(element, true, true) ;
    }

    /* public Part getMainDocumentPart() {
        return templateMainDocumentPart;
    } */

    static int docId =  1000;
    public Drawing createDrawing(byte[] imageByte, Part part, double cmSize) throws Exception {

         
        //https://www.tabnine.com/code/java/classes/org.DocXAPI.openpackaging.parts.WordprocessingML.BinaryPartAbstractImage
        BinaryPartAbstractImage imagePart ;
        if (part == null) {
            imagePart = BinaryPartAbstractImage.createImagePart(templatePackage, imageByte);
        } else { //Part is needed to create image for header and footer part !
            imagePart = BinaryPartAbstractImage.createImagePart(templatePackage, part, imageByte);
        }

        Inline inline;        
        //int docId = ranNumber.nextInt(10000) + 1000; 
        if (cmSize <= 0 ) 
            inline = imagePart.createImageInline(null, null, docId, docId+1, false) ;
        else {
            long twips = Math.round(cmSize * 556) ;            
            inline = imagePart.createImageInline(null, null, docId, docId+1, twips, false);
        } 
        docId = docId + 2 ;  

        Drawing drawing = objectFactory.createDrawing();
        drawing.getAnchorOrInline().add(inline);
        
        //CTCaption txt = objectFactory.createCTCaption() ;
        //txt.setName("hello");
        //drawing.getAnchorOrInline().add(txt) ;
        return drawing;
    }

    public P createImageParagraph(Object image, Part part, double cmSize, JcEnumeration justification) throws Exception {
                                                       //566 twips = 1cm.
        
        //http://prog3.com/sbdm/blog/isea533/article/details/49806637

        P para = objectFactory.createP();
        //R run = objectFactory.createR();
        R run = createRun();
        
        justifyParagraph(para, justification);

        Drawing drawing = null;
        if (image instanceof byte[]) { 
            drawing = createDrawing((byte[])image, part, cmSize) ; 
        }

        if (image instanceof ImageField) {
            ImageField imageField = (ImageField) image;
            //System.out.println(imageField.getSize()) ;
            drawing = createDrawing(imageField.getImage(), part, imageField.getSize()) ; 
        }

        if (drawing != null) {
            run.getContent().add(drawing);
            para.getContent().add(run);
        }
        
        return para;
    } 

    public  P createTextParagraph(Object obj, JcEnumeration justification) throws Exception {

        P para = objectFactory.createP();
        //R run = objectFactory.createR();
        R run = createRun() ;

        justifyParagraph(para, justification) ;
            
        Text text = objectFactory.createText();
        text.setSpace("preserve");
        if (obj instanceof java.lang.String) { 
            text.setValue((String) obj);
        }
            
        if (obj instanceof TextField) {
                    
            TextField txField = (TextField) obj;
            text.setValue(txField.getText());
            
            RPr runPR = objectFactory.createRPr() ;
            runPR = getTextFieldProperties(runPR, txField);
            if (runPR != null) run.setRPr(runPR);
        }
        
        run.getContent().add(text) ;
        para.getContent().add(run);
        return para;
    }

    public void createSimpleParagraph(Object obj, 
                                      Part documentPart,  
                                      JcEnumeration justification) throws Exception {

        P p = objectFactory.createP() ;
        boolean addNothing = true;
        if (obj instanceof java.lang.String || obj instanceof TextField) {
            p = createTextParagraph(obj, justification) ;
            addNothing = false;
        }

        if (obj instanceof byte[] || obj instanceof ImageField) {
            p = createImageParagraph(obj, documentPart, 0, justification) ;
            addNothing = false;
        } 

   /*     if (obj instanceof ImageField) {
            ImageField imageField = (ImageField) obj ;
            p = createImageParagraph(imageField, documentPart, imageField.getSize(), justification) ;
            addNothing = false;
        } */ 

        if (addNothing) return ;

        if (documentPart.getClass().getSimpleName().equalsIgnoreCase("MainDocumentPart"))
            ((MainDocumentPart) documentPart).getContent().add(p) ;

        if (documentPart.getClass().getSimpleName().equalsIgnoreCase("HeaderPart")) 
            ((HeaderPart) documentPart).getContent().add(p) ;
            
        if (documentPart.getClass().getSimpleName().equalsIgnoreCase("FooterPart")) 
            ((FooterPart) documentPart).getContent().add(p) ;
    }

    
    public void createJustifiedParagraph(List<List<Object>> objectList, Part documentPart) throws Exception {
        //NOTE : The list in objectList is expected to be an array of 2 object, the 1st will left justified 
        //       and the 2nd right justified.  the 1st can be a null object to create a paragraph 
        //       with right justified paragraph based on Tab Alignment

        int writableWidthTwips = templatePackage.getDocumentModel()
                                                .getSections().get(0)
                                                .getPageDimensions()
                                                .getWritableWidthTwips();
        int colNum = 2;
        int cellWidth = (writableWidthTwips/colNum) ;
        Tbl table = TblFactory.createTable(objectList.size(), colNum, cellWidth);
        table.setTblPr(new TblPr());  
        CTBorder border = new CTBorder();  
        border.setColor("auto");  
        border.setSz(BigInteger.valueOf(0));  
        border.setSpace(BigInteger.valueOf(0));  
        border.setVal(STBorder.NONE);   
    
        TblBorders borders = new TblBorders();  
        borders.setBottom(border);  
        borders.setLeft(border);  
        borders.setRight(border);  
        borders.setTop(border);  
        borders.setInsideH(border);  
        borders.setInsideV(border);  
        table.getTblPr().setTblBorders(borders);

        int r = -1 ;
        
        for (List<Object> list : objectList) {   //Each element can be 1 line.
        
            int elementIndex = 0 ;
            r ++;
            Tr tRow = (Tr) table.getContent().get(r) ; 

            P[] paragraph = new P[2] ;
            paragraph[0] = objectFactory.createP();
            paragraph[1] = objectFactory.createP();
            
            R[] run = new R[2] ; 
            run[0] =  createRun() ;  
            run[1] =  createRun() ;  

            Boolean isPageNo = false;
            for (Object obj : list) {  //List of complex object
                elementIndex ++ ; 
               
                if (obj != null) {

                    if (obj instanceof TextField) {
                        Text newText = objectFactory.createText();
                        TextField txField = (TextField) obj;
                        newText.setValue(txField.getText());
                        newText.setSpace("preserve");
                        if (txField.getText().equalsIgnoreCase(Constant.PAGENUM))  
                            isPageNo = true;   
                        RPr runPR = objectFactory.createRPr() ;
                        runPR = getTextFieldProperties(runPR, txField); 
                        run[elementIndex-1].setRPr(runPR);
                        run[elementIndex-1].getContent().add(newText) ;
                    } 

                    if (obj instanceof ImageField) {
                        ImageField imageField = (ImageField) obj;
                        Drawing drawing = createDrawing(imageField.getImage(), documentPart, imageField.getSize()) ;
                        run[elementIndex-1].getContent().add(drawing) ; 
                    }
                                            
                    if (obj instanceof java.lang.String) {
                            
                        Text text = objectFactory.createText(); 
                        text.setValue((String)obj);
                     // text.setSpace("preserve");
                        if (text.getValue().equalsIgnoreCase(Constant.PAGENUM))    
                            isPageNo = true ;
                        run[elementIndex-1].getContent().add(text);
                    } 
                    
                    if (obj instanceof byte[]) {
                            Drawing drawing = createDrawing((byte[])obj, documentPart, 0) ;
                            run[elementIndex-1].getContent().add(drawing);
                          //  paragraph.getContent().add(run);
                    } // object is byte[] 
                  

                    Tc tCol ;
                    if (elementIndex % 2 == 0) {
                        justifyParagraph(paragraph[elementIndex-1], JcEnumeration.RIGHT) ;
                        tCol = (Tc) tRow.getContent().get(1);                        
                    } else {
                        justifyParagraph(paragraph[elementIndex-1], JcEnumeration.LEFT) ;
                        tCol = (Tc) tRow.getContent().get(0);
                    }
                  
                    tCol.getContent().clear();
                    if (isPageNo) {
                        P pageNo = createPageNumberParagaph() ;
                        justifyParagraph(pageNo, JcEnumeration.RIGHT);
                        tCol.getContent().add(pageNo) ;  
                    } else {
                        paragraph[elementIndex-1].getContent().add(run[elementIndex-1]);
                        tCol.getContent().add(paragraph[elementIndex-1]) ;
                    }
                  //  System.out.println("*** " + elementIndex +" ==>" + XmlUtils.marshaltoString(paragraph)) ;

                } // if object is not null
                
            }  //Inner For object
            
            
        } // Outer For object

        if (documentPart.getClass().getSimpleName().equalsIgnoreCase("MainDocumentPart"))
            ((MainDocumentPart) documentPart).getContent().add(table) ;

        if (documentPart.getClass().getSimpleName().equalsIgnoreCase("HeaderPart")) 
            ((HeaderPart) documentPart).getContent().add(table) ;
        
        if (documentPart.getClass().getSimpleName().equalsIgnoreCase("FooterPart"))
            ((FooterPart) documentPart).getContent().add(table) ; 
    
    }    
       
      

    public CTSimpleField createSimpleField(String val) {
		
		CTSimpleField field = new CTSimpleField();
		field.setInstr(val);
	    return field;
	}

    public void createComplexField(P p, String instrText) {
		
		
		// Create object for r
	   // R r = objectFactory.createR(); 
        R r = createRun() ;
	    p.getContent().add(r); 
        
        // Create object for fldChar (wrapped in JAXBElement) 
        FldChar fldchar = objectFactory.createFldChar(); 
        JAXBElement<org.docx4j.wml.FldChar> fldcharWrapped = objectFactory.createRFldChar(fldchar); 
        r.getContent().add( fldcharWrapped); 
        fldchar.setFldCharType(org.docx4j.wml.STFldCharType.BEGIN);
        
        // Create object for instrText (wrapped in JAXBElement) 
        Text text = objectFactory.createText(); 
        JAXBElement<org.docx4j.wml.Text> textWrapped = objectFactory.createRInstrText(text); 
        r.getContent().add( textWrapped); 
        text.setValue( instrText); 
        text.setSpace( "preserve"); 	

        // Create object for fldChar (wrapped in JAXBElement) 
        fldchar = objectFactory.createFldChar(); 
        fldcharWrapped = objectFactory.createRFldChar(fldchar); 
        r.getContent().add( fldcharWrapped); 
        fldchar.setFldCharType(org.docx4j.wml.STFldCharType.END);
		
	}

    //https://java.hotexamples.com/examples/org.DocXAPI.wml/FldChar/-/java-fldchar-class-examples.html
    //https://github-wiki-see.page/m/plutext/DocXAPI-ImportXHTML/wiki/Page-Footer 
    public P createPageNumberParagaph() throws Exception {

   /*      String pageNumbering = 
            "<w:p xmlns:w=\"http://schemas.openxmlformats.org/wordprocessingml/2006/main\">"
            + "<w:r><w:rPr><w:noProof/></w:rPr><w:t xml:space=\"preserve\">Page </w:t></w:r>"
            + "<w:r><w:rPr><w:noProof/></w:rPr><w:fldChar w:fldCharType=\"begin\"/></w:r>"
            + "<w:r><w:rPr><w:noProof/></w:rPr><w:instrText xml:space=\"preserve\"> PAGE  \\* Arabic  \\* MERGEFORMAT </w:instrText></w:r>"
            + "<w:r><w:rPr><w:noProof/></w:rPr><w:fldChar w:fldCharType=\"separate\"/></w:r>"
            + "<w:r><w:rPr><w:noProof/></w:rPr><w:fldChar w:fldCharType=\"end\"/></w:r>"
            + "<w:r><w:rPr><w:noProof/></w:rPr><w:t xml:space=\"preserve\"> of </w:t></w:r>"
            + "<w:r><w:rPr><w:noProof/></w:rPr><w:fldChar w:fldCharType=\"begin\"/></w:r>"
            + "<w:r><w:rPr><w:noProof/></w:rPr><w:instrText xml:space=\"preserve\"> NUMPAGES  \\* Arabic  \\* MERGEFORMAT </w:instrText></w:r>"
            + "<w:r><w:rPr><w:noProof/></w:rPr><w:fldChar w:fldCharType=\"separate\"/></w:r>"
            + "<w:r><w:rPr><w:noProof/></w:rPr><w:fldChar w:fldCharType=\"end\"/></w:r></w:p>" ;
        // BooleanDefaultTrue TRUE  = new BooleanDefaultTrue();
        //R run = objectFactory.createR() ; 
        Object pageNumberingObject = XmlUtils.unmarshalString(pageNumbering) ;
        return (P) pageNumberingObject ;
        
        */
        P paragraph = objectFactory.createP() ;
        paragraph.getContent().add(createSimpleField(" PAGE \\* MERGEFORMAT "));
        //createComplexField(paragraph, " PAGE \\* MERGEFORMAT ");
        return paragraph ;        
    }


    public R createRun() {
        RPr runProperties = objectFactory.createRPr();
        changeFontToArial(runProperties);
      /*  changeFontSize(runProperties,11) ; */ 
        R run = objectFactory.createR() ;
        run.setRPr(runProperties);
        return run ;
    }

    public Text createText(String value)
    {
        Text t = objectFactory.createText();
        t.setValue(value);
        t.setSpace("preserve");
        return t;
    }
       
    public void clearHeaderFooter() {
            
            // Remove from sectPr
            SectPrFinder finder = new SectPrFinder(templateMainDocumentPart);
            new TraversalUtil(templateMainDocumentPart.getContent(), finder);
            for (SectPr sectPr : finder.getOrderedSectPrList()) {
                sectPr.getEGHdrFtrReferences().clear() ;                 
            }
            
            // Remove rels
            List<Relationship> hfRels = new ArrayList<Relationship>(); 
            for (Relationship rel : templateMainDocumentPart.getRelationshipsPart().getRelationships().getRelationship() ) {
                
                if (rel.getType().equals(Namespaces.HEADER) || 
                    rel.getType().equals(Namespaces.FOOTER)) {
                    hfRels.add(rel);
                }
            }

            for (Relationship rel : hfRels ) {
                templateMainDocumentPart.getRelationshipsPart().removeRelationship(rel);
            }
            
    }

    private int headerFooterCounter = 100 ;
    public FooterPart createFtrPart(HdrFtrRef location) throws Exception {

        FooterPart footerPart = new FooterPart(new PartName("/word/header" + (headerFooterCounter++) + ".xml"));
        footerPart.setPackage(templatePackage);
        
        Relationship relationship = templateMainDocumentPart.addTargetPart(footerPart); 
        
        FooterReference footerReference = objectFactory.createFooterReference();
        footerReference.setId(relationship.getId());
        footerReference.setType(location);

        List<SectionWrapper> sections = templatePackage.getDocumentModel().getSections();
        // There is always a section wrapper, but it might not contain a sectPr
        SectPr sectPr = sections.get(sections.size() - 1).getSectPr();
        if (sectPr == null) {
            sectPr = objectFactory.createSectPr();        
            sections.get(sections.size() - 1).setSectPr(sectPr);
        }       
        sectPr.getEGHdrFtrReferences().add(footerReference);// add header or footer
        templateMainDocumentPart.addObject(sectPr);
        return footerPart ;
    }

    
    public HeaderPart createHdrPart(HdrFtrRef location) throws Exception {
        
        HeaderPart headerPart = new HeaderPart(new PartName("/word/header" + (headerFooterCounter++) + ".xml"));
        headerPart.setPackage(templatePackage);
        
        Relationship relationship = templateMainDocumentPart.addTargetPart(headerPart); 
        
        HeaderReference headerReference = objectFactory.createHeaderReference();
        headerReference.setId(relationship.getId());
        headerReference.setType(location);

        List<SectionWrapper> sections = templatePackage.getDocumentModel().getSections();
        // There is always a section wrapper, but it might not contain a sectPr
        SectPr sectPr = sections.get(sections.size() - 1).getSectPr();
        if (sectPr == null) {
            sectPr = objectFactory.createSectPr();        
            sections.get(sections.size() - 1).setSectPr(sectPr);
        }       
        sectPr.getEGHdrFtrReferences().add(headerReference);// add header or footer
        templateMainDocumentPart.addObject(sectPr); 
        return headerPart ;
    }


    public FooterPart getFtrPart(WordprocessingMLPackage wordPackage, HdrFtrRef type) {
      
        List<SectionWrapper> sectionWrappers = wordPackage.getDocumentModel().getSections();
		
		for (SectionWrapper sw : sectionWrappers) {
			HeaderFooterPolicy hfp = sw.getHeaderFooterPolicy();
            if (type == HdrFtrRef.FIRST) 
                return hfp.getFirstFooter() ;

            if (type == HdrFtrRef.DEFAULT) 
                return hfp.getDefaultFooter() ;

            if (type == HdrFtrRef.EVEN) 
                return hfp.getEvenFooter();
        }

        return null;            
    } 

    public HeaderPart getHdrPart(WordprocessingMLPackage wordPackage, HdrFtrRef type) {
      
        List<SectionWrapper> sectionWrappers = wordPackage.getDocumentModel().getSections();
		
		for (SectionWrapper sw : sectionWrappers) {
			HeaderFooterPolicy hfp = sw.getHeaderFooterPolicy();
            if (type == HdrFtrRef.FIRST) 
                return hfp.getFirstHeader() ;

            if (type == HdrFtrRef.DEFAULT) 
                return hfp.getDefaultHeader() ;

            if (type == HdrFtrRef.EVEN) 
                return hfp.getEvenHeader();
        }

        return null;            
    } 

    public void createStdHeader(List<Object> objects, HdrFtrRef location, JcEnumeration justification) throws Exception {
                                                                //location : FIRST, EVEN, DEFAULT etc
        HeaderPart headerPart = createHdrPart(location) ;

        for (Object obj : objects) {
                if (obj instanceof java.lang.String || obj instanceof TextField) {     
                    headerPart.getContent().add(createTextParagraph(obj, justification)) ;
                } else {
                    if (obj instanceof byte[] || obj instanceof ImageField) { 
                        headerPart.getContent().add(createImageParagraph(obj, headerPart, 0, justification ));
                    } else {
                       // System.out.println("NOTHING !!! " + obj.getClass().getSimpleName());
                    }
                }
        } 
    }


    public void createStdFooter(List<Object> objects, HdrFtrRef location, JcEnumeration justification) throws Exception {

        //clearHeaderFooter();
        FooterPart footerPart = createFtrPart(location) ;
         
        for (Object obj : objects) {
            if (obj instanceof java.lang.String || obj instanceof TextField) {      
                footerPart.getContent().add(createTextParagraph(obj, justification)) ;
            } else {
                if (obj instanceof byte[] || obj instanceof ImageField) {
                    footerPart.getContent().add(createImageParagraph(obj, footerPart, 0, justification));
                } 
            }
        } 
        
        P pageParagraph = createPageNumberParagaph() ;
        footerPart.getContent().add(pageParagraph) ;


    }

    public void createHeaderObject(List<List<Object>> imageOrTextList, HdrFtrRef location) throws Exception {

        HeaderPart hdrPart = getHdrPart(templatePackage, location) ;
        if (hdrPart == null) {
            hdrPart = createHdrPart(location) ;
        }
        createJustifiedParagraph(imageOrTextList, hdrPart);

    }

    public void createFooterObject(List<List<Object>> imageOrTextList, HdrFtrRef location) throws Exception {

        FooterPart ftrPart = getFtrPart(templatePackage, location) ;
        if (ftrPart == null) {
            ftrPart = createFtrPart(location) ;
        }
        createJustifiedParagraph(imageOrTextList, ftrPart);

    }

    //https://iprofs.wordpress.com/2012/11/19/adding-layout-to-your-DocXAPI-generated-word-documents-part-2/

    public void changeGlobalFont(String fontName) {
        
        Styles styles = templateMainDocumentPart.getStyleDefinitionsPart ().getJaxbElement ();
         
         for (Style s : styles.getStyle ()) {
            //System.out.println("Name:" + s.getName().getVal());
        
            if (s.getName ().getVal ().equalsIgnoreCase("Normal")) {

                RPr rpr = s.getRPr ();
                if (rpr == null) {
                    rpr = objectFactory.createRPr ();
                    s.setRPr (rpr);
                }
                RFonts rf = rpr.getRFonts ();
                if (rf == null) {
                    rf = objectFactory.createRFonts ();
                    rpr.setRFonts (rf);
                }
                // This is where you set your font name.
                rf.setAscii (fontName);
            }
        }
    }

    public void changeFontToArial(RPr runProperties) {
        RFonts runFont = new RFonts();
        runFont.setAscii("Arial");
        runFont.setHAnsi("Arial");
        runProperties.setRFonts(runFont);
    }


    public void changeFontSize(RPr runProperties, int fontSize) {
        HpsMeasure size = new HpsMeasure();
        size.setVal(BigInteger.valueOf(fontSize));
        runProperties.setSz(size);
    }
    
    public  List<Object> getAllElementFromObject(Object obj, Class<?> toSearch) {
        List<Object> result = new ArrayList<>();
        if (obj instanceof JAXBElement) obj = ((JAXBElement<?>) obj).getValue();

        //DELETE
        //System.out.println(obj.getClass().getSimpleName()) ;
      
        if (obj.getClass().equals(toSearch))
            result.add(obj);
        else if (obj instanceof ContentAccessor) {
            List<?> children = ((ContentAccessor) obj).getContent();
            for (Object child : children) {
                result.addAll(getAllElementFromObject(child, toSearch));
            }
        }
        return result;
    } 


    public void substituteField(Part part, String fieldName, Object replacementObject) throws Exception {

        List<Object> paragraphs = getAllElementFromObject(part, P.class) ;
        
        for(Object p : paragraphs) {
            P paragraph = (P) p ;
            List<Object> texts = getAllElementFromObject(paragraph, Text.class) ;
            
            for (Object txt : texts) {
                Text text = (Text) txt ;
                List<String> tokens = textTokenizer(text.getValue()) ; 
                RPr runRPr = objectFactory.createRPr();
                Object r = XmlUtils.unwrap(text.getParent());
                if (r instanceof R) {
                    R run = (R) r ; 
                    run.getContent().remove(text) ; 
                    runRPr = run.getRPr();
                }  
                
                //log.info("Text {} - Tokenized paragraphs {}", text.getValue(), tokens) ;
                for (String t : tokens) {
                    
                    //System.out.println("Field name=" + fieldName + " - "  + t) ;

                    //R newRun = objectFactory.createR() ;
                    R newRun = createRun() ;
                    Text newText = objectFactory.createText();
                    newText.setSpace("preserve") ;  
                    if (t.contains(fieldName)) {
                        //System.out.print ("2") ;
                        
                        if (replacementObject instanceof TextField) {
                          //  System.out.print ("3") ;
                            TextField textField = (TextField) replacementObject ;
                            RPr newRPr = getTextFieldProperties(runRPr, textField) ;
                            newText.setValue(textField.getText()); 
                            newRun.getContent().add(newText) ; 
                            newRun.setRPr(newRPr);  
                        } 

                        if (replacementObject instanceof ImageField) {

                            ImageField  imageField = (ImageField) replacementObject ;
                            newRun.getContent().add(
                                   createDrawing(imageField.getImage(), part, imageField.getSize())) ;
                        }
                    } else {
                       
                        newText.setValue(t) ; 
                        newRun.setRPr(runRPr) ;
                        newRun.getContent().add(newText) ;
                    } 
                     
                      
                    paragraph.getContent().add(newRun) ;
                                        
                }

            }
        
        }

}



    public void replaceField(String fieldName, Object obj) throws Exception {
        //http://vixmemon.blogspot.com/2013/04/DocXAPI-replace-text-placeholders-with.html
        
        substituteField(templateMainDocumentPart,fieldName, obj) ;

        RelationshipsPart rP = templateMainDocumentPart.getRelationshipsPart();
    
        List<Relationship> rS = rP.getRelationshipsByType(Namespaces.HEADER);
        for (Relationship relationship : rS) {
            HeaderPart headerPart = (HeaderPart) rP.getPart(relationship.getId()) ;
                  
            substituteField(headerPart, fieldName, obj) ;
        }

        rS = rP.getRelationshipsByType(Namespaces.FOOTER);
        for (Relationship relationship : rS) {
            FooterPart footerPart = (FooterPart) rP.getPart(relationship.getId()) ;
            substituteField(footerPart, fieldName, obj) ;
        }
      
    }

    

    public void replaceBookmark(String bookmarkName, Object obj, double cmSize) throws Exception {
                                   
        Document wmlDoc = templateMainDocumentPart.getJaxbElement();
        Body body = wmlDoc.getBody();
        List<Object> paragraphs = body.getContent();                    
              
        RangeFinder rt = new RangeFinder("CTBookmark", "CTMarkupRange");  // Extract bookmarks and create bookmark cursor
        new TraversalUtil(paragraphs, rt);
        
        // traverse bookmarks
        for (CTBookmark bm : rt.getStarts()) {
        
            // Here you can operate on a single bookmark, or you can use a map to process all bookmarks
            if (bm.getName().equals(bookmarkName)) {

                P p = (P) (bm.getParent());
                //R run = objectFactory.createR();
                R run = createRun() ;
                if (obj instanceof byte[]) {
                    Drawing drawing = createDrawing((byte[])obj, templateMainDocumentPart, cmSize);
                    run.getContent().add(drawing);
                }   
               
                if (obj instanceof java.lang.String) {
                    Text txt = objectFactory.createText() ;
                    txt.setValue((String)obj);
                    run.getContent().add(txt);
                }
               
                p.getContent().add(run);
            }
        }
    }

    public  void replaceTable(String[] placeholders, List<Map<String, String>> textToAdd, MainDocumentPart mainDocumentPart) throws Exception, JAXBException {
        
        if (placeholders[0] != null) {
            
            List<Object> tables = getAllElementFromObject(mainDocumentPart, Tbl.class);       
            if (tables.size() == 0) {
                //TODO: log info
                return ;
            } 
            Tbl tempTable = getTemplateTable(tables, placeholders[0]);        
            List<Object> rows = getAllElementFromObject(tempTable, Tr.class);
            
    //      if (rows.size() == 1) { //careful only tables with 1 row are considered here
                Tr templateRow = (Tr) rows.get(rows.size()-1);
                for (Map<String, String> replacements : textToAdd) {              
                    addRowToTable(tempTable, templateRow, replacements);
                }
                assert tempTable != null;
                tempTable.getContent().remove(templateRow);
    //      } 
        }
        
    }

    private  void addRowToTable(Tbl reviewTable, Tr templateRow, Map<String, String> replacements) {
        Tr workingRow = XmlUtils.deepCopy(templateRow);
        List<?> textElements = getAllElementFromObject(workingRow, Text.class);
        for (Object object : textElements) {
            Text text = (Text) object;
            String replacementValue = replacements.get(text.getValue());
            if (replacementValue != null)
                text.setValue(replacementValue);
        }
        reviewTable.getContent().add(workingRow);
    }

    private  Tbl getTemplateTable(List<Object> tables, String templateKey) {
        
        for (Object tbl : tables) {
            List<?> textElements = getAllElementFromObject(tbl, Text.class);            
            for (Object text : textElements) {
                Text textElement = (Text) text;
                
                if (textElement.getValue() != null && textElement.getValue().equals(templateKey))
                    return (Tbl) tbl;
            }
        }
        return null;
    }

    public TextField createTextField(String text) {
        TextField txField =  new TextField() ;
        return txField.setText(text) ;
        //return txField ;

    }

    public RPr getTextFieldProperties(RPr runPR, TextField textField) {
        
        RPr cloneRPr = objectFactory.createRPr() ; 
        if (runPR != null) {
            cloneRPr = XmlUtils.deepCopy(runPR) ;
        }
        
        BooleanDefaultTrue turnOn = new BooleanDefaultTrue();
        Highlight highlight = objectFactory.createHighlight() ;
        Color fontColor = objectFactory.createColor();
        
        if (textField.isBold()) {
            cloneRPr.setB(turnOn);
        }      
        if (textField.isItalics()) {
            cloneRPr.setI(turnOn);
        }
        if (textField.isUnderline()) {
            U underline = objectFactory.createU();
            underline.setVal(UnderlineEnumeration.SINGLE) ;
            cloneRPr.setU(underline);
        }      

        if (textField.isHighlightedGreen()) {
            //color.setVal("green");
            highlight.setVal("green");
            cloneRPr.setHighlight(highlight);
        }
        if (textField.isHighlightedYellow()) {
            highlight.setVal("yellow");
            cloneRPr.setHighlight(highlight);
        }

        if (textField.isRedFont()) {
            fontColor.setVal("red");    
            cloneRPr.setColor(fontColor);
        }

        if (textField.isBlueFont()) {
            fontColor.setVal("blue");    
            cloneRPr.setColor(fontColor);
        }
        
       // System.out.println(runPR + " value:" + this.getText()) ;
        return cloneRPr ;

    }

    public ImageField createImageField(byte[] imageByte) {
        ImageField imageField = new ImageField();
        return imageField.setImage(imageByte) ;
    }

    public List<Object> createParagraphList(Object leftJustified, Object rightJustified) {

        List<Object> line = new ArrayList<>();
        line.add(leftJustified) ;
        line.add(rightJustified) ;
        return line;
    }

    public void mergeHeaderFooter (String srcFileName) throws Exception {

        List<HdrFtrRef> hdrFtrRefList = 
                        Arrays.asList(HdrFtrRef.FIRST, HdrFtrRef.EVEN, HdrFtrRef.DEFAULT);
        List<String> headerFooterList = Arrays.asList("HEADER", "FOOTER") ;

        WordprocessingMLPackage sourcePackage = WordprocessingMLPackage.load(new File(srcFileName));
            
        for (String hf : headerFooterList) {

                HeaderPart sourceHeaderPart = null ;
                FooterPart sourceFooterPart = null ;
                HeaderPart newHeaderPart = null;
                FooterPart newFooterPart = null ;
                RelationshipsPart sourceRelationshipsPart = null ;
                for (HdrFtrRef hfRef : hdrFtrRefList) {
                    
                    if (hf.equals("HEADER")){
                        sourceHeaderPart = getHdrPart(sourcePackage,hfRef) ;
                        if (sourceHeaderPart != null) {
                            newHeaderPart = createHdrPart(hfRef) ;
                            try {
                                sourceRelationshipsPart = sourceHeaderPart.getRelationshipsPart();
                            } catch (Exception e) {}
                        }           
                        
                    } else {
                        if (hf.equals("FOOTER")){
                            sourceFooterPart =  getFtrPart(sourcePackage,hfRef) ;
                            if (sourceFooterPart != null) {
                                newFooterPart = createFtrPart(hfRef) ;
                                try {
                                    sourceRelationshipsPart = sourceFooterPart.getRelationshipsPart();
                                } catch (Exception e) {}
                                
                            }           
                            
                        }
                    }

                    //The following lines of code is to look for image in the header/footer
                   // RelationshipsPart sourceRelationshipsPart = sourceHeaderFooterPart.getRelationshipsPart();
                    if (sourceRelationshipsPart != null) { 
                        List<Relationship> sourceRelationships = sourceRelationshipsPart.getRelationships().getRelationship();
                        for (Relationship r : sourceRelationships) {
                            //System.out.print(" | Part type:" + r.getType() + " | Extension:"); 
                            if (r.getType().contains("image")) {
                                Part part = sourceRelationshipsPart.getPart(r);
                                if (part != null) {

                                    //The following if is to ensure that the image is loaded into memory for part
                                    //All IMAGE (PNG, JPEG or GIF are classified under BinaryPart.
                                    //Hence just need to load once using getBuffer
                                    if (part instanceof BinaryPart) {   //MUST BE the first check !
                                        //System.out.print("BIN->") ;
                                        ((BinaryPart)part).getBuffer();
                                    }
                                    
                                    if (part instanceof ImagePngPart) {
                                        //System.out.print("PNG") ;
                                    //    ((ImagePngPart) part).getBytes() ;
                                    }
                    
                                    if (part instanceof ImageJpegPart) {
                                        //System.out.print("JPG") ;
                                        //((ImageJpegPart) part).getBytes() ;
                                    }

                                    if (part instanceof ImageGifPart) {
                                       // System.out.print("GIF") ;
                                        //((ImageGifPart) part).getBytes() ;
                                    }
                                    
                        
                                    part.setPackage(templatePackage);
                                    Relationship imgRelation = templateMainDocumentPart.addTargetPart(part,
                                                                    AddPartBehaviour.RENAME_IF_NAME_EXISTS) ;
                                    r.setId(imgRelation.getId());

                                    //The following is VERY IMPORTANT
                                    if (hf.equals("HEADER")) 
                                        BinaryPartAbstractImage.createImagePart(templatePackage,
                                                                    newHeaderPart, ((BinaryPartAbstractImage)part).getBytes());
                                    else 
                                        BinaryPartAbstractImage.createImagePart(templatePackage,
                                                                    newFooterPart, ((BinaryPartAbstractImage)part).getBytes());
                                    

                                } //if image part not null

                            } // End looking for image
                        } // End for image relationship 
                    }  //End of checking for sourceRelationship 

                    if (sourceHeaderPart != null) {
                        for (Object obj : sourceHeaderPart.getContent()) {
                            newHeaderPart.getContent().add(obj) ;
                        }
                    }

                    if (sourceFooterPart != null) {
                        for (Object obj : sourceFooterPart.getContent()) {
                            newFooterPart.getContent().add(obj) ;
                        }
                    }
                    
                } //For HdrFtrRef
        } //For header and footer looping
    }
    

    public void mailMerge() throws Exception {

        //-------------- Implementation logic 
        
        //Define the header & footer
        if (Boolean.TRUE.equals(nothingToMerge())) {
            log.debug("Nothing to mail merge.") ;
            return ;
        }
        
        if (headerFooters.size() > 0 ) {
            for (Map.Entry<String, HeaderFooter> hdr : headerFooters.entrySet()) { 
                HeaderFooter hf = hdr.getValue() ;  
                if (hdr.getKey().contains("HEADER")) {
                    createHeaderObject(hf.getParagraphs(), hf.getHdrFtrRef());
                    //createStdHeader(hf.getParagraphs(), hf.getHdrFtrRef(), hf.getJustification());
                }   
                if (hdr.getKey().contains("FOOTER")) {
                    createFooterObject(hf.getParagraphs(), hf.getHdrFtrRef());
                    //createStdFooter(hf.getParagraphs(), hf.getHdrFtrRef(), hf.getJustification());
                }  
                
            }
        } 

        //Table Merge field                       
        if (mergeField.tables.size() > 0) { 
            log.debug("Merging tables.");
            String[] colNames = null ; 
            for (Map.Entry<String, List<Map<String, String>>> tableList :mergeField.tables.entrySet()) {
                
                List <Map<String, String>> dataList = tableList.getValue();  
                if (dataList.size() > 0) {
                    
                    Map<String, String>dataMap = dataList.get(0) ;
                    colNames = new String[dataMap.size()] ;
                    int i = -1 ;
                    for (Map.Entry<String, String> dataRow : dataMap.entrySet()) {
                        i++ ;
                        colNames[i] = dataRow.getKey() ;
                    }    

                }
                replaceTable (colNames, dataList, templateMainDocumentPart) ;
            }

        }
         
        
        // Image Merge Field
        String fieldName;
        if (mergeField.imageVariables.size() > 0) {
            log.debug("Merging images.");
            for (Map.Entry<String, ImageField> imageField : mergeField.imageVariables.entrySet()) {
                fieldName = imageField.getKey() ;              
                replaceField(fieldName, imageField.getValue());
            }
        } 

        //Text Merge Fields
        if (mergeField.textVariables.size() > 0) {
            log.debug("Merging text.");
            for (Map.Entry<String, TextField> textFields : mergeField.textVariables.entrySet()) {
                fieldName = textFields.getKey() ;
                if (textFields.getValue().isEncrypted() && textFields.getValue().getText().startsWith("!")) {
                    try {
                        String text = Crypto.aesBase64Decrypt(textFields.getValue().getText().getBytes()) ;
                        TextField txtField = createTextField(text) ;
                        String fmt = textFields.getValue().getFormats().replace("_", "") ;
                        txtField.setFormats(fmt);
                        replaceField(fieldName, txtField);
                    } catch (Exception e) {}
                } else 
                    replaceField(fieldName, textFields.getValue());
            }
        }
    
    } 

/*
    public List<String> backupTextTokenizer(String text) {

        List<String> textToken = new ArrayList<>();
        String fieldName = "" ;
        String textValue = "" ;

        boolean isFieldName = false; 
        int lastFieldIndex = 0 ;

        for (int i=0;i<text.length();i++) {
            if (text.charAt(i) == '{') isFieldName = true;
            
            String t = new StringBuilder().append(text.charAt(i)).toString() ;
            if (isFieldName) {                 
                fieldName = fieldName.concat(t) ;
                if (textValue.length() > 0) {
                    textToken.add(textValue) ;
                    textValue = "";
                }
               
            } else {
                textValue = textValue.concat(t) ;
            }

            if (text.charAt(i) == '}') {
                textToken.add(fieldName) ;
                isFieldName = false ;
                fieldName = "";
                lastFieldIndex = i;  
            }
        }

        if (lastFieldIndex < text.length() && textValue.length() > 0 ) {  //got residual text to add to array after for loop
            textToken.add(textValue) ;
        }

        return textToken;

    }
*/

    public List<String> textTokenizer(String text) {

        List<String> textToken = new ArrayList<>();
        String fieldName = "" ;
        String textValue = "" ;

        boolean isFieldName = false; 
        int lastFieldIndex = 0 ;

        for (int i=0;i<text.length();i++) {

            String t = new StringBuilder().append(text.charAt(i)).toString() ;
            if (text.charAt(i) == '{') isFieldName = true;
            
            if (isFieldName) {                 
                fieldName = fieldName.concat(t) ;
                if (textValue.length() > 0) {
                    textToken.add(textValue) ;
                    textValue = "";
                }
            } else {
                if (text.charAt(i) != '{' && text.charAt(i) != '}') //<= double protection to get rid of { and }
                    textValue = textValue.concat(t) ;
            }
                    
            if (text.charAt(i) == '}') {                
                textToken.add(fieldName) ;
                isFieldName = false ;
                fieldName = "";
                lastFieldIndex = i;  
            }
        }

        if (lastFieldIndex < text.length() && textValue.length() > 0 ) {  //got residual text to add to array after for loop
            textToken.add(textValue) ;
        }

        return textToken;

    }

    
    public void findAndReplace(String target,  String replacement) throws Exception {

        List<Object> texts = getAllElementFromObject(templateMainDocumentPart, Text.class) ;
        for (Object text : texts) {
            Text txt = (Text) text ;
            String replacedString = txt.getValue().replaceAll(target, replacement) ;
            txt.setValue(replacedString); 
        }
    }

    public void insertParagraph(String target, P paragraph) {

        List<Object> paragraphs = getAllElementFromObject(templateMainDocumentPart, P.class) ;
        int pIndex = 0 ;
        for (Object para : paragraphs) {
            P p = (P) para ;
            pIndex ++ ;
            List<Object> texts = getAllElementFromObject(p, Text.class) ;
            for (Object text : texts) {
                Text txt = (Text) text ;
                
                if (txt.getValue().contains(target)) {
                     
                    templateMainDocumentPart.getContent().add(pIndex,paragraph) ;
                }
            }
        }
    }

    public byte[] createQRCode(String text, int width, int height) throws Exception {
        
        String charSet = "UTF-8" ;
        Map<EncodeHintType, Object>  hintMap = new EnumMap<>(EncodeHintType.class);
        hintMap.put(EncodeHintType.ERROR_CORRECTION, ErrorCorrectionLevel.Q);
        hintMap.put(EncodeHintType.MARGIN, -1);   //NOTE:change -1 to 1 if qr code cannot be scanned)

        BitMatrix matrix = 
            new MultiFormatWriter().encode(new String(text.getBytes(charSet), charSet),
                                           BarcodeFormat.QR_CODE, width, height, hintMap);

        ByteArrayOutputStream bos = new ByteArrayOutputStream(); 
        MatrixToImageWriter.writeToStream(matrix, "png", bos);
    
        return bos.toByteArray() ; 
            
    }

    public static void main(String[] args) throws Exception {
        
    }

}

