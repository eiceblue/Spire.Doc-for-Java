import com.spire.doc.*;
import com.spire.doc.documents.*;
import com.spire.doc.fields.*;

public class resetShapeSize {
    public static void main(String[] args) {
        String input="data/shapes.docx";
        String output="output/resetShapeSize.docx";

        //load a document
        Document doc = new Document();
        doc.loadFromFile(input);

        //get the first section and the first paragraph that contains the shape
        Section section = doc.getSections().get(0);
        Paragraph para = section.getParagraphs().get(0);

        //get the second shape and reset the width and height for the shape
        ShapeObject shape = (ShapeObject)para.getChildObjects().get(1) ;
        shape.setWidth (200);
        shape.setHeight(200);

        //save the document
        doc.saveToFile(output, FileFormat.Docx);
    }
}
