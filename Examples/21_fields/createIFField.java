import com.spire.doc.*;
import com.spire.doc.documents.*;
import com.spire.doc.fields.*;
import com.spire.doc.interfaces.IParagraphBase;

public class createIFField {
    public static void main(String[] args) throws Exception {
        //Create Word document.
        Document document = new Document();

        //Add a new section.
        Section section = document.addSection();

        //Add a new paragraph.
        Paragraph paragraph = section.addParagraph();

        // Define a method of creating an IF Field.
        CreateIfField(document, paragraph);

        //Define merged data.
        String[] fieldName = { "Count" };
        String[] fieldValue = { "2" };

        //Merge data into the IF Field.
        document.getMailMerge().execute(fieldName, fieldValue);

        //Update all fields in the document.
        document.isUpdateFields(true);

        //Save to file.
        String result = "output/CreateAnIFField.docx";
        document.saveToFile(result, FileFormat.Docx_2013);
    }
    //Create the IF Field like:{IF { MERGEFIELD Count } > "100" "Thanks" " The minimum order is 100 units "}
    static void CreateIfField(Document document, Paragraph paragraph)
    {
        IfField ifField = new IfField(document);
        ifField.setType(FieldType.Field_If);
        ifField.setCode("IF");

        paragraph.getItems().add(ifField);
        paragraph.appendField("Count", FieldType.Field_Merge_Field);
        paragraph.appendText(" > ");
        paragraph.appendText("\"100\" ");
        paragraph.appendText("\"Thanks\" ");
        paragraph.appendText("\"The minimum order is 100 units\"");

        IParagraphBase endPara = document.createParagraphItem(ParagraphItemType.Field_Mark);
        FieldMark end=(FieldMark)endPara;
        end.setType(FieldMarkType.Field_End);
        paragraph.getItems().add(end);
        ifField.setEnd(end) ;
    }
}
