import com.spire.doc.*;
import com.spire.doc.documents.*;
import com.spire.doc.fields.*;

public class alignShape {
    public static void main(String[] args) {
        String input = "data/shapes.docx";
        String output = "output/alignShape.docx";

        //load Document
        Document doc = new Document();
        doc.loadFromFile(input);

        //get the shapes
        Section section = doc.getSections().get(0);
        for (int i=0;i<section.getParagraphs().getCount();i++)
        {
            Paragraph para = section.getParagraphs().get(i);
            for (int j=0;j<para.getChildObjects().getCount();j++)
            {
                DocumentObject obj = para.getChildObjects().get(j);
                if (obj instanceof ShapeObject)
                {
                    //set the horizontal alignment as center
                    ((ShapeObject) obj).setHorizontalAlignment(ShapeHorizontalAlignment.Center);
                }
            }
        }
        //save the document
        doc.saveToFile(output, FileFormat.Docx);
    }
}
