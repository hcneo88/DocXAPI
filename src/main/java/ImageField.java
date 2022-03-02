import java.util.HashMap;
import java.util.Map;
import java.util.Base64;

public class ImageField {

    private String image ;
    private Map<String, Object> properties = new HashMap<>();

    public ImageField setImage(String base64Image) {
        image = base64Image;
        return this;
    }

    public ImageField setImage(byte[] byteImage) {
        image = Base64.getEncoder().encodeToString(byteImage);
        this.setSize(0) ;
        return this;
    }

    public ImageField setSize(double cmSize) {
        properties.put(Constant.IMG_PROP_SIZE, cmSize) ;            
        return this ;
    }

    public ImageField setProperties (Map<String, Object> prop) {
        properties.putAll(prop) ;
        return this;
    }
    public byte[] getImage() {
        return Base64.getDecoder().decode(image) ;
    }

    public double getSize() {
        if (properties.containsKey(Constant.IMG_PROP_SIZE))
            return (double)properties.get(Constant.IMG_PROP_SIZE) ;
        else
            return 0 ;
    }
    
}
