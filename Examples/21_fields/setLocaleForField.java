import com.spire.doc.*;
import com.spire.doc.documents.*;
import com.spire.doc.fields.*;

public class setLocaleForField {
    public static void main(String[] args) {
        //Open a Word document
        Document document = new Document("data/sampleB_2.docx");

        //Get the first section
        Section section = document.getSections().get(0);

        Paragraph par = section.addParagraph();

        //Add a date field
        Field field = par.appendField("DocDate", FieldType.Field_Date);

        //Set the LocaleId for the textrange
        TextRange range = (TextRange) field.getOwnerParagraph().getChildObjects().get(0);
        range.getCharacterFormat().setLocaleIdASCII((short) 1049);

        field.setFieldText("2019-10-10");
        //Update field
        document.isUpdateFields(true);

        String output = "output/setLocaleForField.docx";
        document.saveToFile(output, FileFormat.Docx_2013);
    }
}
