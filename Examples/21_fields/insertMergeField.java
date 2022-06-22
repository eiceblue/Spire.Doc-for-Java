import com.spire.doc.*;
import com.spire.doc.documents.*;
import com.spire.doc.fields.*;

public class insertMergeField {
    public static void main(String[] args) {
        Document document = new Document("data/sampleB_2.docx");

        //Get the first section
        Section section = document.getSections().get(0);

        //Add paragraph
        Paragraph par = section.addParagraph();

        //Add merge field in the paragraph
        MergeField field = (MergeField) par.appendField("MyFieldName", FieldType.Field_Merge_Field);

        //Save to file
        String output = "output/insertMergeField.docx";
        document.saveToFile(output, FileFormat.Docx_2013);
    }
}
