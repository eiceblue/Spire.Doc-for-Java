import com.spire.doc.*;
import com.spire.doc.documents.*;
import com.spire.doc.fields.*;

public class rotateShape {
    public static void main(String[] args) {
        String input="data/shapes.docx";
        String output="output/rotateShape.docx";

        //load a document
        Document doc = new Document();
        doc.loadFromFile(input);

        //get the first section
        Section section = doc.getSections().get(0);

        //traverse the word document and set the shape rotation as 20
        for (int i=0;i<section.getParagraphs().getCount();i++)
        {
            Paragraph para = section.getParagraphs().get(i);
            for (int j=0;j<para.getChildObjects().getCount();j++)
            {
                DocumentObject obj =para.getChildObjects().get(j);
                if (obj instanceof ShapeObject)
                {
               ((ShapeObject) obj).setRotation(20.0);
                }
            }
        }
        //save the document
        doc.saveToFile(output, FileFormat.Docx);
    }
}
