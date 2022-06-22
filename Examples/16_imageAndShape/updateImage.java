import com.spire.doc.*;
import com.spire.doc.documents.*;
import com.spire.doc.fields.*;
import java.util.ArrayList;

public class updateImage {
    public static void main(String[] args) {
        String input1="data/imageTemplate.docx";
        String input2="data/e-iceblue.png";
        String output="output/updateImage.docx";

        //load a document
        Document doc = new Document();
        doc.loadFromFile(input1);

        //create a list
        ArrayList<DocumentObject> pictures = new ArrayList<DocumentObject>();

        //get all pictures in the Word document
        for (int i=0; i< doc.getSections().getCount();i++)
        {
            Section sec= doc.getSections().get(i);
            for (int p=0; p< sec.getParagraphs().getCount();p++)
            {
                Paragraph para =sec.getParagraphs().get(p);
                //get all pictures in the Word document
                for (int o=0; o< para.getChildObjects().getCount();o++)
                {
                    DocumentObject docObj= para.getChildObjects().get(o);
                    if (docObj.getDocumentObjectType() == DocumentObjectType.Picture)
                    {
                        pictures.add(docObj);
                    }
                }
            }
        }
        //replace the first picture with a new image file
        DocPicture picture = (DocPicture)pictures.get(0) ;
        picture.loadImage(input2);

        //save the document
        doc.saveToFile(output, FileFormat.Docx);
    }
}
