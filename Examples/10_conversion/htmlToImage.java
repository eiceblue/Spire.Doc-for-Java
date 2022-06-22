import com.spire.doc.*;
import com.spire.doc.documents.*;

import javax.imageio.ImageIO;
import java.awt.image.BufferedImage;
import java.io.File;

public class htmlToImage {
    public static void main(String[] args) throws Exception{
        //Create Word document.
        Document document = new Document();

        //Load the file from disk.
        document.loadFromFile("data/Template_HtmlFile1.html", FileFormat.Html, XHTMLValidationType.None);


        //Save to image. You can convert HTML to BMP, JPEG, PNG, GIF, Tiff etc.
        BufferedImage image= document.saveToImages(0, ImageType.Bitmap);
        String result = "output/result-HtmlToImage.png";
        File file= new File(result);
        ImageIO.write(image, "PNG", file);
    }
}
