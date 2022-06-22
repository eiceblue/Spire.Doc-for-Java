import com.spire.doc.*;
import com.spire.doc.documents.*;
import com.spire.doc.fields.*;

public class removeField {
    public static void main(String[] args) {
        //Open a Word document
        Document document = new Document("data/ifFieldSample.docx");

        //Get the first field
        Field field = document.getFields().get(0);

        //Get the paragraph of the field
        Paragraph par = field.getOwnerParagraph();
        //Get the index of the  field
        int index = par.getChildObjects().indexOf(field);
        //Remove if field via index
        par.getChildObjects().removeAt(index);

        //Save doc file
        String output = "output/removeField.docx";
        document.saveToFile(output, FileFormat.Docx_2013);
    }
}
