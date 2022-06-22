import com.spire.doc.*;
import com.spire.doc.documents.*;
import com.spire.doc.fields.*;

public class setCultureForField {
    public static void main(String[] args) {
        //Create a document
        Document document = new Document();

        //Create a section
        Section section = document.addSection();

        //Add paragraph
        Paragraph paragraph = section.addParagraph();

        //Add textRnage for paragraph
        paragraph.appendText("Add Date Field: ");

        //Add date field1
        Field field1 = paragraph.appendField("Date1", FieldType.Field_Date);
        field1.setCode("DATE  \\@" + "\"yyyy\\MM\\dd\"");

        //Add new paragraph
        Paragraph newParagraph = section.addParagraph();

        //Add textRnage for paragraph
        newParagraph.appendText("Add Date Field with setting French Culture: ");

        //Add date field2
        Field field2 = newParagraph.appendField("\"\\@\"dd MMMM yyyy", FieldType.Field_Date);

        //Setting Field with setting French Culture
        field2.getCharacterFormat().setLocaleIdASCII((short) 1036);

        //Update fields
        document.isUpdateFields(true);

        //Save the document
        String output = "output/setCultureForField.docx";
        document.saveToFile(output, FileFormat.Docx_2013);
    }
}
