import com.spire.doc.*;
import com.spire.doc.documents.*;
import com.spire.doc.fields.*;

public class removeShape {
    public static void main(String[] args) {
        String input="data/shapes.docx";
        String output="output/removeShape.docx";

        //load a word document
        Document doc = new Document();
        doc.loadFromFile(input);

        //get first section
        Section section = doc.getSections().get(0);

        //traverse all the child objects of paragraph
        for (int j= 0; j<section.getParagraphs().getCount();j++)
        {
            Paragraph para = section.getParagraphs().get(j);
            for (int i = 0; i < para.getChildObjects().getCount(); i++)
            {
                //if the child objects is shape object
                if (para.getChildObjects().get(i) instanceof ShapeObject)
                {
                    //remove the shape object
                    para.getChildObjects().removeAt(i);
                }
            }
        }
        //Save the document
        doc.saveToFile(output, FileFormat.Docx);
    }
}
