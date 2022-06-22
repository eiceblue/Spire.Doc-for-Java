import com.spire.doc.*;
import com.spire.doc.documents.*;
import com.spire.doc.fields.*;

public class addTCField {
    public static void main(String[] args) {
        //Create Word document
        Document document = new Document();

        //Add a new section
        Section section = document.addSection();

        //Add a new paragraph
        Paragraph paragraph = section.addParagraph();

        //Add TC field in the paragraph
        Field field = paragraph.appendField("TC", FieldType.Field_TOC_Entry);
        field.setCode("TC " + "\"Entry Text\"" + " \\f" + " t");

        //Save to file
        document.saveToFile("output/addTCField.docx", FileFormat.Docx);
    }
}
