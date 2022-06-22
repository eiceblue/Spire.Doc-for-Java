import com.spire.doc.*;
import com.spire.doc.documents.*;
import com.spire.doc.fields.*;
import java.util.ArrayList;

public class replaceImageWithText {
    public static void main(String[] args) {
        String input="data/imageTemplate.docx";
        String output="output/replaceImageWithText.docx";

        //load a document
        Document doc = new Document();
        doc.loadFromFile(input);

        //replace all pictures with texts
        int j = 1;
        for (int i=0; i< doc.getSections().getCount();i++)
        {
            Section sec= doc.getSections().get(i);
            for (int p=0; p< sec.getParagraphs().getCount();p++)
            {
                Paragraph para =sec.getParagraphs().get(p);
                ArrayList<DocumentObject> pictures = new ArrayList<DocumentObject>();

                //get all pictures in the Word document
                for (int o=0; o< para.getChildObjects().getCount();o++)
                {
                    DocumentObject docObj= para.getChildObjects().get(o);
                    if (docObj.getDocumentObjectType() == DocumentObjectType.Picture)
                    {
                        pictures.add(docObj);
                    }
                }
                //replace pitures with the text "Here was image {image index}"
                for (int m=0; m<pictures.size();m++)
                {
                    DocumentObject pic= pictures.get(m);
                    int index = para.getChildObjects().indexOf(pic);
                    TextRange range = new TextRange(doc);
                    range.setText(String.format("Here was image-%d", j));
                    para.getChildObjects().insert(index, range);
                    para.getChildObjects().remove(pic);
                    j++;
                }
            }
        }
        //save the document
        doc.saveToFile(output, FileFormat.Docx);
    }
}
