import com.spire.doc.*;
import com.spire.doc.documents.*;

import javax.imageio.*;
import java.awt.image.*;
import java.io.*;

public class toImageWithHighResolution {
    public static void main(String[] args) throws IOException {
        String input = "data/imageResolutionTemplate.docx";
        String output = "output/toImageWithHighResolution.png";

        //load Word document
        Document document= new Document();
        document.loadFromFile(input);
        //set the image resolution and save to image
        BufferedImage[] image= document.saveToImages(0, 1, ImageType.Bitmap, 300, 300);
        File file= new File(output);
        ImageIO.write(image[0], "PNG", file);
    }
}
