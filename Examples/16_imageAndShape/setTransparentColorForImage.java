import com.spire.doc.*;
import com.spire.doc.documents.*;
import com.spire.doc.fields.*;
import java.awt.*;

public class setTransparentColorForImage {
    public static void main(String[] args) {
        String input="data/imageTemplate.docx";
        String output="output/setTransparentColorForImage.docx";

        //load a document
        Document doc = new Document();
        doc.loadFromFile(input);

        //get the first paragraph in the first section
        Paragraph paragraph = doc.getSections().get(0).getParagraphs().get(0);

        //set the blue color of the image(s) in the paragraph to transperant
        for (int i=0;i<paragraph.getChildObjects().getCount();i++)
        {
            DocumentObject obj = paragraph.getChildObjects().get(i);
            if (obj instanceof DocPicture)
            {
                DocPicture picture = (DocPicture)obj;
                picture.setTransparentColor(Color.BLUE);
            }
        }
        //save the document
        doc.saveToFile(output, FileFormat.Docx);
    }
}
