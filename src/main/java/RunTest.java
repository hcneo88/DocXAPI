import java.io.File;
import java.io.FileInputStream;
import java.io.PrintStream;
import java.security.MessageDigest;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

import org.apache.commons.io.FileUtils;
import org.apache.commons.lang3.time.StopWatch;
import org.docx4j.openpackaging.parts.DrawingML.Drawing;
import org.docx4j.wml.JcEnumeration;
import org.docx4j.wml.P;
import org.docx4j.wml.Text;


public class RunTest {
    
    public static String getFileChecksum(File file) throws Exception {

            MessageDigest digest = MessageDigest.getInstance("MD5");
            
            //Get file input stream for reading the file content
            FileInputStream fis = new FileInputStream(file);
            
            //Create byte array to read data in chunks
            byte[] byteArray = new byte[1024];
            int bytesCount = 0; 
                
            //Read file data and update in message digest
            while ((bytesCount = fis.read(byteArray)) != -1) {
                digest.update(byteArray, 0, bytesCount);
            };
            
            //close the stream; We don't need it now.
            fis.close();
            
            //Get the hash's bytes
            byte[] bytes = digest.digest();
            
            //This bytes[] has bytes in decimal format;
            //Convert it to hexadecimal format
            StringBuilder sb = new StringBuilder();
            for(int i=0; i< bytes.length ;i++)
            {
                sb.append(Integer.toString((bytes[i] & 0xff) + 0x100, 16).substring(1));
            }
            
            //return complete hash
            return sb.toString();
    }

    public static Boolean compareFileContent (String testId, String firstFile, String secondFile) throws Exception {
        
        File file1  = new File(firstFile) ;
        File file2 = new File(secondFile) ;

        String firstChecksum  = getFileChecksum(file1) ;
        String secondChecksum = getFileChecksum(file2) ; 

        Boolean testResult = firstChecksum.equalsIgnoreCase(secondChecksum) ;
        System.out.println(testId + " [" + firstChecksum + ", " 
                                                     + secondChecksum + "] - Pass ? "  
                                                     + testResult.toString().toUpperCase()) ;
        return  testResult;

    }

    static DocXAPI wordAPI = new DocXAPI(); 
    static PrintStream originalOut    = System.out;
    
    static String templatePath ;
    static String expectedPath ;
    static String resultPath ;
    static StopWatch stopWatch ;

    public static void setUpTest(String testResultLogFile, boolean deleteFolder) throws Exception {

        templatePath = RunTest.class.getResource("/template").toURI().getPath(); 
        expectedPath = RunTest.class.getResource("/test/expected").toURI().getPath() ;
        resultPath   = RunTest.class.getResource("/test/result").toURI().getPath() ;
            
        if (deleteFolder) {
            File resultFolder = new File(resultPath) ; 
            if (resultFolder.exists()) {
                FileUtils.cleanDirectory(resultFolder) ;
            } 
            resultFolder.mkdir();
        } 

        PrintStream ps = new PrintStream(resultPath + "/" + testResultLogFile) ; 
        System.setOut(ps);
        stopWatch = new StopWatch();
        stopWatch.start() ;
        
    }
    
    public static void generateResult(String testId, String file2Compare ) throws Exception {

        stopWatch.stop();
        System.setOut(originalOut);
        Boolean passFail = true ;
            
        if (file2Compare.endsWith(".docx")) {
            wordAPI.loadTemplate(resultPath + "/" + file2Compare) ;
            List<Object> resultParagraph = 
                wordAPI.getAllElementFromObject(wordAPI.getTemplateMainDocumentPart(), P.class);
            
            wordAPI.loadTemplate(expectedPath + "/" + file2Compare) ;
            List<Object> expectedParagraph = 
                wordAPI.getAllElementFromObject(wordAPI.getTemplateMainDocumentPart(), P.class);
       
            if (expectedParagraph.size() == resultParagraph.size()) {
                for (int i=0;i<expectedParagraph.size();i++) {
                    P ePara = (P) expectedParagraph.get(i) ;
                    P rPara = (P) resultParagraph.get(i) ;
                    List<Object> eTextList = wordAPI.getAllElementFromObject(ePara, Text.class) ;
                    List<Object> rTextList = wordAPI.getAllElementFromObject(rPara, Text.class) ;
                    for (int k=0;k<eTextList.size();k++) {
                        Text eTxt = (Text) eTextList.get(k) ;
                        Text rTxt = (Text) rTextList.get(k) ;
                        passFail = rTxt.getValue().equals(eTxt.getValue());
                            //System.out.println(k+":"+ rTxt.getValue() + ":" + eTxt.getValue() + " " + passFail); 
                            //passFail = false;
                        if (! passFail) break;

                    } 

                    if (! passFail) break;

                    List<Object> eImageList = wordAPI.getAllElementFromObject(ePara, Drawing.class) ;
                    List<Object> rImageList = wordAPI.getAllElementFromObject(rPara, Drawing.class) ;
                    for (int j=0;j<eImageList.size();j++) {
                        Drawing eImage = (Drawing) eImageList.get(j) ;
                        Drawing rImage = (Drawing) rImageList.get(j) ;
                        passFail = rImage.getPartName().equals(eImage.getPartName());
                        if (! passFail) break;
                    }
                    if (! passFail) break;
                    
                }
            } else passFail = false ;
 
        } else {  //NOT a docx file.  Direct compare.
            File resultFile = new File(resultPath + "/" + file2Compare);
            File expectedResultFile = new File(expectedPath + "/" + file2Compare);
            passFail = FileUtils.contentEquals(resultFile, expectedResultFile);
        }

        System.out.println(testId + " (" + file2Compare +  ") took " + stopWatch.getTime() +  " ms to execute. Result - Pass ? "  + passFail.toString().toUpperCase()) ;
    }

    public static void textTokenizerTestCase(String templateFilename) throws Exception {
        
        try {
            if (templateFilename != null) wordAPI.loadTemplate(templatePath + "/" + templateFilename) ; 
            List<String> tokenList ;
            String txt = "{fieldname} is going to rain." ;
            tokenList = wordAPI.textTokenizer(txt) ;
            System.out.println("==============") ;
            tokenList = wordAPI.textTokenizer(txt) ;
            for (String l : tokenList) {
                System.out.println(l) ;
            }
            System.out.println("==============") ;
            
            txt = "He is going to {location}" ;
            tokenList = wordAPI.textTokenizer(txt) ;
            for (String l : tokenList) {
                System.out.println(l) ;
            }
            System.out.println("Size=" + tokenList.size()) ;
            System.out.println("==============") ;

            txt = "He wakes up at {time} this morning." ;
            tokenList = wordAPI.textTokenizer(txt) ;
            for (String l : tokenList) {
                System.out.println(l) ;
            } 
            
            System.out.println("==============") ;

            txt = "{name} is coming here at {time} tommmorrow morning." ;
            tokenList = wordAPI.textTokenizer(txt) ;
            for (String l : tokenList) {
                System.out.println(l) ;
            } 
            
            System.out.println("==============") ;

            txt = "{name} said he is waiting for us at {time} on {day}" ;
            tokenList = wordAPI.textTokenizer(txt) ;
            for (String l : tokenList) {
                System.out.println(l) ;
            } 
            
            System.out.println("==============") ;

            txt = "{name} said he is waiting.  I told {name} to wait at {location} instead." ;
            tokenList = wordAPI.textTokenizer(txt) ;
            for (String l : tokenList) {
                System.out.println(l) ;
            } 
                    
            System.out.println("==============") ;


            txt = "{photo}" ;
            tokenList = wordAPI.textTokenizer(txt) ;
            for (String l : tokenList) {
                System.out.println(l) ;
            } 
            System.out.println("Size=" + tokenList.size()) ;        
            System.out.println("==============") ;
        } catch (Exception e) {
            System.out.println(e.toString()) ;
        } 
        
}

public static void simpleParagraphTestCase(String templateFilename) throws Exception {
   
    
    wordAPI.loadTemplate(templatePath + "/" + templateFilename);
    byte[] imageByte = wordAPI.loadImage("d:/data/work/java/DocXAPI/ncs-logo.png") ;
    wordAPI.createSimpleParagraph("LEFT", wordAPI.getTemplateMainDocumentPart(), JcEnumeration.LEFT);
    wordAPI.createSimpleParagraph("MIDDLE", wordAPI.getTemplateMainDocumentPart(), JcEnumeration.CENTER);
    wordAPI.createSimpleParagraph("RIGHT", wordAPI.getTemplateMainDocumentPart(), JcEnumeration.RIGHT);

    wordAPI.createSimpleParagraph(wordAPI.createTextField("hello").italics(), wordAPI.getTemplateMainDocumentPart(), JcEnumeration.LEFT);
    wordAPI.createSimpleParagraph(wordAPI.createTextField("hello").bold(), wordAPI.getTemplateMainDocumentPart(), JcEnumeration.CENTER);
    wordAPI.createSimpleParagraph(wordAPI.createTextField("hello").bold().italics(), wordAPI.getTemplateMainDocumentPart(), JcEnumeration.RIGHT);

    wordAPI.createSimpleParagraph(imageByte, wordAPI.getTemplateMainDocumentPart(), JcEnumeration.LEFT);
    wordAPI.createSimpleParagraph(imageByte, wordAPI.getTemplateMainDocumentPart(), JcEnumeration.CENTER);
    wordAPI.createSimpleParagraph(imageByte, wordAPI.getTemplateMainDocumentPart(), JcEnumeration.RIGHT);

    wordAPI.createSimpleParagraph(wordAPI.createImageField(imageByte).setSize(1), wordAPI.getTemplateMainDocumentPart(), JcEnumeration.LEFT);
    wordAPI.createSimpleParagraph(wordAPI.createImageField(imageByte).setSize(2), wordAPI.getTemplateMainDocumentPart(), JcEnumeration.CENTER);
    wordAPI.createSimpleParagraph(wordAPI.createImageField(imageByte).setSize(3), wordAPI.getTemplateMainDocumentPart(), JcEnumeration.RIGHT);
  
    wordAPI.saveDoc(resultPath + "/" + templateFilename);
    wordAPI.saveAsPDF(resultPath + "/" + templateFilename, null) ;
}

public static void useLetterHeadTestCase(String templateFilename,  String letterHeadname) throws Exception {
   
    wordAPI.loadTemplate(templatePath + "/" + templateFilename) ;
    wordAPI.mergeHeaderFooter(templatePath + "/LetterHead/" + letterHeadname);

        String longText = "Lorem Ipsum is simply dummy text of the printing and typesetting industry."
                + " Lorem Ipsum has been the industry's standard dummy text ever since the 1500s, when"
                + " an unknown printer took a galley of type and scrambled it to make a type specimen"
                + " book. It has survived not only five centuries, but also the leap into electronic" 
                + "typesetting, remaining essentially unchanged." ;
        
        Map<String, TextField> mappings = new HashMap<>();        
        mappings.put("Name",      wordAPI.createTextField("Zion Tan").italics().encrypt(true));
        mappings.put("Line1",     wordAPI.createTextField("10 Windsor Park"));
        mappings.put("Line2",     wordAPI.createTextField("#09-600").red());        
        mappings.put("Line3",     wordAPI.createTextField("Singapore 129143"));
        mappings.put("Reference", wordAPI.createTextField("T23094-1123112345678")) ;
        mappings.put("Date",      wordAPI.createTextField("20 Jan 2022")) ;
        mappings.put("LongText",  wordAPI.createTextField(longText));
        wordAPI.setTextMergeFields(mappings) ;

        Map<String, ImageField> images = new HashMap<>();
        byte[] imageByte = wordAPI.createQRCode("T23094-1123112345678", 50, 50);
        images.put("QR", wordAPI.createImageField(imageByte).setSize(0)) ;
        wordAPI.setImageMergeFields(images) ;

        wordAPI.createSimpleParagraph(wordAPI.createTextField("The Russian is attacking Ukraine").italics(), wordAPI.getTemplateMainDocumentPart(), JcEnumeration.LEFT);
        
        //Create Object Paragraph with left and right justification

        imageByte = wordAPI.loadImage(templatePath + "/images/eagle.jpg") ;
        List<List<Object>> list = new ArrayList<>() ;
        List<Object> line = wordAPI.createParagraphList(
                                wordAPI.createTextField("Feature A - Insert split justified objects : via TextField & ImageField").bold().italics().underline(),
                                wordAPI.createImageField(imageByte).setSize(0)) ; 
        list.add(line) ;

        imageByte = wordAPI.loadImage(templatePath + "/images/lion.jpg") ;
        line = wordAPI.createParagraphList(imageByte, 
                                        "Feature B - Insert split justified objects : String and byte[]") ; 
        list.add(line) ; 

        line = wordAPI.createParagraphList("Feature C - Insert split justified objects with only left portion", 
                                         null ) ; 
        list.add(line) ; 

        line = wordAPI.createParagraphList(null, "Feature D - right portion") ; 
        list.add(line) ; 
        wordAPI.createJustifiedParagraph(list, wordAPI.getTemplateMainDocumentPart());

        wordAPI.mailMerge();
        wordAPI.saveDoc(resultPath + "/" + templateFilename); 
        wordAPI.saveAsPDF(resultPath + "/" + templateFilename,  null);
}

    public static void main(String[] args) throws Exception {

        /*
        setUpTest("textTokenizerTestCase.log", true) ;
        textTokenizerTestCase(null) ;
        generateResult("id-textTokenizer","textTokenizerTestCase.log") ;

        setUpTest("simpleParagraphTestCase.log",false) ;
        simpleParagraphTestCase("testTemplate.docx");
        generateResult("id-simpleParagraph","simpleParagraphTestCase.log") ;
    */
    
        setUpTest("useLetterHeadTestCase.log", false) ;
        useLetterHeadTestCase("emptyTemplate.docx", "LetterHead.docx");
        generateResult("id-useLetterHeaderTestCase","emptyTemplate.docx");
        

    }

}
