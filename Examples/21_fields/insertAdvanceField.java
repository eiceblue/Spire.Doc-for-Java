import com.spire.doc.*;
import com.spire.doc.documents.Paragraph;
import com.spire.doc.fields.Field;

public class insertAdvanceField {
    public static void main(String[] args) {
        //Open a Word document
        String input="data/BlankTemplate.docx";
        Document document = new Document(input);

        //Get the first section
        Section section = document.getSections().get(0);

        Paragraph par = section.addParagraph();

        //Add advance field
        Field field = par.appendField("Field", FieldType.Field_Advance);

        //Add field code
        field.setCode("ADVANCE \\d 10 \\l 10 \\r 10 \\u 0 \\x 100 \\y 100 ");

        //Update field
        document.isUpdateFields(true);

        //Save to file
        String result="output/insertAdvanceField.docx";
        document.saveToFile(result, FileFormat.Docx);
    }
}
