import java.nio.charset.StandardCharsets;
import java.security.MessageDigest;
import java.security.SecureRandom;
import java.security.spec.KeySpec;
import java.util.Base64;
import java.util.Random;

import javax.crypto.Cipher;
import javax.crypto.SecretKey;
import javax.crypto.SecretKeyFactory;
import javax.crypto.spec.SecretKeySpec;


import javax.crypto.spec.GCMParameterSpec;
import javax.crypto.spec.PBEKeySpec;

public class Crypto {

    public static String SYSTEM_SEED = "JavaAln'tC0ffee" ;   //PLEASE DO NOT Change
    public static String SYSTEM_SALT = "YeARtwo022" ;        //PLEASE DO NOT Change
    public static int SYSTEM_SECRET_LENGTH = 16; 

    public static final int GCM_IV_LENGTH = 12;
    public static final int GCM_TAG_LENGTH = 16;

    Crypto() {}

    static Random random = new Random();

    public static String getRandString(int length) {
        int leftLimit = 48;     // numeral '0'
        int rightLimit = 122;   // letter 'z'
        
        if (length <= 0) length = 8;
        String generatedString = random.ints(leftLimit, rightLimit + 1)
        .filter(i -> (i <= 57 || i >= 65) && (i <= 90 || i >= 97))
        .limit(length)
        .collect(StringBuilder::new, StringBuilder::appendCodePoint, StringBuilder::append)
        .toString();

        return generatedString ;

    }

    public static String bytesToHex(byte[] hash) {
        StringBuilder hexString = new StringBuilder(2 * hash.length);
        for (int i = 0; i < hash.length; i++) {
            String hex = Integer.toHexString(0xff & hash[i]);
            if(hex.length() == 1) {
                hexString.append('0');
            }
            hexString.append(hex);
        }
        return hexString.toString();
    }

    public static String createHash(String inputString) throws Exception {
        
        final MessageDigest digest = MessageDigest.getInstance("SHA-256");
        final byte[] hashbytes = digest.digest(
        inputString.getBytes(StandardCharsets.UTF_8));
        return bytesToHex(hashbytes);
    
    } 
    
    public static String createSecret(String seed) throws Exception {
        
        String    hash = createHash(seed + SYSTEM_SEED) ; 
        if (SYSTEM_SECRET_LENGTH >= hash.length()) {
            hash = hash.concat(SYSTEM_SEED) ;
        }       
        return hash.substring(2, SYSTEM_SECRET_LENGTH) ;

    }

    public static SecretKey getKeyFromSecret(String secret) throws Exception {
    
        SecretKeyFactory factory = SecretKeyFactory.getInstance("PBKDF2WithHmacSHA256");
        KeySpec spec = new PBEKeySpec(secret.toCharArray(), SYSTEM_SALT.getBytes(), 65536, 256);
        SecretKey encryptionKey = new SecretKeySpec(factory.generateSecret(spec).getEncoded(), "AES");
        return encryptionKey;
    } ;

    public static byte[] createIV() {

        byte[] IV = new byte[GCM_IV_LENGTH];
        SecureRandom random = new SecureRandom();
        random.nextBytes(IV);
        return IV ;
    }

    public static byte[] aesEncrypt(byte[] plainByte, String secret, byte[] IV) throws Exception {
        
        Cipher cipher = Cipher.getInstance("AES/GCM/NoPadding");

        SecretKey key = getKeyFromSecret(secret) ;
        SecretKeySpec keySpec = new SecretKeySpec(key.getEncoded(), "AES");

        GCMParameterSpec gcmParameterSpec = new GCMParameterSpec(GCM_TAG_LENGTH * 8, IV);
        cipher.init(Cipher.ENCRYPT_MODE, keySpec, gcmParameterSpec);
        
        // Perform Encryption
        byte[] cipherText = cipher.doFinal(plainByte);
        
        return cipherText;

    }
    
    
    public static String aesDecrypt(byte[] cipherByte, String secret, byte[] IV) throws Exception {
        
        Cipher cipher = Cipher.getInstance("AES/GCM/NoPadding");

        SecretKey key = getKeyFromSecret(secret) ;
        SecretKeySpec keySpec = new SecretKeySpec(key.getEncoded(), "AES");

        GCMParameterSpec gcmParameterSpec = new GCMParameterSpec(GCM_TAG_LENGTH * 8, IV);
        cipher.init(Cipher.DECRYPT_MODE, keySpec, gcmParameterSpec);
        
        // Perform Encryption
        byte[] plainText = cipher.doFinal(cipherByte);
        
        return new String(plainText);

    }

    public static String aesBase64Encrypt(byte[] plainByte, String seed) throws Exception {
        
        String secret = seed + Constant.SYSTEM_SEED ;
        byte[] IV = Crypto.createIV() ;
        byte[] cipherText = Crypto.aesEncrypt(plainByte, secret, IV) ;

        String base64Seed =  Base64.getEncoder().encodeToString(seed.getBytes()) ;
        String base64IV = Base64.getEncoder().encodeToString(IV) ;
        String base64CipherText = Base64.getEncoder().encodeToString(cipherText) ;

        //Prepend ! to indicate that this is an encrypted string
        String base64EncodedCipher = "!" + base64IV + "." + getRandString(8) + base64Seed + "." + base64CipherText ;
        
        return base64EncodedCipher;

    }

    public static String aesBase64Decrypt(byte[] base64EncodedByte ) throws Exception {
        
        String cipherString = new String(base64EncodedByte) ;
        String[] cipherText = cipherString.split("\\.") ;
        String ivPart = cipherText[0].substring(1) ;                //remove the ! marker 

        String seedPart = cipherText[1] ;
        seedPart = seedPart.substring(8, seedPart.length()) ;

        String cipherPart = cipherText[2] ;
        
        byte[] base64DecodedCipher = Base64.getDecoder().decode(cipherPart.getBytes()) ;
        byte[] base64DecodedIV = Base64.getDecoder().decode(ivPart.getBytes()) ;
        byte[] base64DecodedSeed = Base64.getDecoder().decode(seedPart.getBytes());
        
        String key = new String(base64DecodedSeed) + Constant.SYSTEM_SEED ; 
        String plainText = aesDecrypt(base64DecodedCipher, 
                                      key,base64DecodedIV) ;
        
        return plainText;
    }

    public static void runCryptoTestCase() throws Exception {

        String plainText = "Covid-19 cases is trending up." ;
        String secret = Crypto.createSecret("Singapore") ;
  
        //Crypto crypto = new Crypto();
   /*     for (int i=0;i<10;i++) {
            //System.out.println(crypto.getRandString(15));
            String rnd = Crypto.getRandString(16) ;
            System.out.println("rnd=" + rnd + " secret=" + Crypto.createSecret(rnd)) ;
        } 

        
        byte[] IV = createIV();
        byte[] cipherText = Crypto.aesEncrypt(plainText.getBytes(), secret, IV) ;
        String clearText = Crypto.aesDecrypt(cipherText, secret, IV) ;
      //  System.out.println("Plain Text=" + plainText + " secret=" + secret + " IV=" + new String(IV)) ;
      //  System.out.println("Encrypted=" + new String(cipherText)) ;
      //  System.out.println("Decrypted=" + clearText) ;
*/
        String base64Text = Crypto.aesBase64Encrypt(plainText.getBytes(), secret) ;
        String decryptedText = Crypto.aesBase64Decrypt(base64Text.getBytes()) ;
        

    }
    
}
