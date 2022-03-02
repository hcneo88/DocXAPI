public class TextField {
    private String value ;
    private String formats = "_";
    
    public TextField setText(String text) {
        value = text ;
        return this ;
    } 

    public TextField setFormats (String fmt) {
        formats = formats.concat(fmt) ;
        return this ;
    }

    public String getText() {
        return value ;
    }

    public String getFormats() {
        return formats;
    }
    
    public boolean isBold() {
        return formats.contains("B") ;
    }

    public boolean isUnderline() {
        return formats.contains("U") ;
    }

    public boolean isItalics() {
        return formats.contains("I") ;
    }

    public boolean isHighlightedYellow() {
        return formats.contains("H:y");
    }

    public boolean isHighlightedGreen() {
        return formats.contains("H:g");
    }

    public boolean isRedFont() {
        return formats.contains("C:r");
    }

    public boolean isBlueFont() {
        return formats.contains("C:b");
    }

    public TextField bold() {
        formats = formats.concat("B") ; 
        return this;           
    }

    public TextField italics() {
        formats = formats.concat("I") ; 
        return this  ;         
    }

    public TextField underline() {
        formats = formats.concat("U") ; 
        return this ;           
    }

    

    public TextField highlightYellow() {
        formats = formats.concat("H:").concat("y") ;
        return this;
    }

    public TextField highlightGreen() {
        formats = formats.concat("H:").concat("g") ;
        return this;
    } 

    
    public TextField red() {
        formats = formats.concat("C:").concat("r") ;
        return this;
    }

    public TextField blue() {
        formats = formats.concat("C:").concat("b") ;
        return this;
    } 

    public TextField encrypt(boolean yesNo) {
        if (yesNo) {
                if (! formats.contains("E")) {
                    String keySeed = Crypto.getRandString(12) ;
                    try {
                        String cipherText = Crypto.aesBase64Encrypt(getText().getBytes(), keySeed) ;
                        setText(cipherText) ;
                        formats = formats.concat("E")    ;
                    } 
                    catch (Exception ex) {} 
                    
                }                 
        } else {

            if (formats.contains("E")) {
                try {
                    String text = Crypto.aesBase64Decrypt(getText().getBytes());
                    setText(text) ;
                    formats = formats.replace("E", "") ;
                } catch (Exception e) {}
            }

        }
        return this ;
    }

    public boolean isEncrypted() {
        return formats.contains("E") ;
    }
}
