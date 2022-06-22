import com.spire.doc.*;
import com.spire.doc.documents.Paragraph;
import com.spire.doc.fields.Field;

public class insertAddressBlockField {
    public static void main(String[] args) {
        //Open a Word document
        String input="data/BlankTemplate.docx";
        Document document = new Document(input);

        //Get the first section
        Section section = document.getSections().get(0);

        Paragraph par = section.addParagraph();

        //Add address block in the paragraph
        Field field = par.appendField("ADDRESSBLOCK", FieldType.Field_Address_Block);

        //Set field code
        field.setCode("ADDRESSBLOCK \\c 1 \\d \\e Test2 \\f Test3 \\l \"Test 4\"");

        //Save to file
        String result="output/insertAddressBlockField.docx";
        document.saveToFile(result, FileFormat.Docx);
    }
}
