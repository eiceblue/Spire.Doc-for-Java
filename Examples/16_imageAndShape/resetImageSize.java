import com.spire.doc.*;
import com.spire.doc.documents.*;
import com.spire.doc.fields.*;

public class resetImageSize {
    public static void main(String[] args) {
        String input="data/imageTemplate.docx";
        String output="output/resetImageSize.docx";

        //load a document
        Document doc = new Document();
        doc.loadFromFile(input);

        //get the first secion
        Section section = doc.getSections().get(0);

        //get the first paragraph
        Paragraph paragraph = section.getParagraphs().get(0);

        //reset the image size of the first paragraph
        for (int i =0; i<paragraph.getChildObjects().getCount();i++)
        {
            DocumentObject docObj= paragraph.getChildObjects().get(i);
            if (docObj instanceof DocPicture)
            {
                DocPicture picture = (DocPicture)docObj;
                picture.setWidth(50f);
                picture.setHeight(50f);
            }
        }
        //save the document
        doc.saveToFile(output, FileFormat.Docx);
    }
}
