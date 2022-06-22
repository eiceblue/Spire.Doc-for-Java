import com.spire.doc.*;
import com.spire.doc.documents.*;
import com.spire.doc.fields.*;

public class insertPageRefField {
    public static void main(String[] args) {
        //Open a Word document
        Document document = new Document("data/pageRef.docx");

        //Get the first section
        Section section = document.getLastSection();

        Paragraph par = section.addParagraph();

        //Add page ref field
        Field field = par.appendField("pageRef", FieldType.Field_Page_Ref);

        //Set field code
        field.setCode("PAGEREF  bookmark1 \\# \"0\" \\* Arabic  \\* MERGEFORMAT");

        //Update field
        document.isUpdateFields(true);

        String output = "output/insertPageRefField.docx";
        document.saveToFile(output, FileFormat.Docx_2013);
    }
}
