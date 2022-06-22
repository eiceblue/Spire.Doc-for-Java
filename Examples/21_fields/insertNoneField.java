import com.spire.doc.*;
import com.spire.doc.documents.*;
import com.spire.doc.fields.*;

public class insertNoneField {
    public static void main(String[] args) {
        Document document = new Document("data/sampleB_2.docx");

        //Get the first section
        Section section = document.getSections().get(0);

        //Add paragraph
        Paragraph par = section.addParagraph();

        //Add a none field
        Field field = par.appendField("", FieldType.Field_None);

        //Save to file
        String output = "output/insertNoneField.docx";
        document.saveToFile(output, FileFormat.Docx_2013);
    }
}
