import com.spire.doc.*;
import com.spire.doc.documents.*;
import com.spire.doc.fields.*;

public class executeConditionalField {

    public static void main(String[] args) throws Exception {

        Document doc = new Document();
        //Add a new section
        Section section = doc.addSection();
        //Add a new paragraph for a section
        Paragraph paragraph = section.addParagraph();

        CreateIFField1(doc, paragraph);
        paragraph = section.addParagraph();
        CreateIFField2(doc, paragraph);

        String[] fieldName = { "Count","Age" };
        String[] fieldValue = { "2","30" };

        doc.getMailMerge().execute(fieldName, fieldValue);
        doc.isUpdateFields(true);

        doc.saveToFile("sample.docx", FileFormat.Docx);

        String result = "output/ExecuteConditionalField_out.docx";
        doc.saveToFile(result, FileFormat.Docx);

    }
    private static void CreateIFField1(Document document, Paragraph paragraph)
    {
        IfField ifField = new IfField(document);
        ifField.setType(FieldType.Field_If);
        ifField.setCode("IF ");
        paragraph.getItems().add(ifField);

        paragraph.appendField("Count", FieldType.Field_Merge_Field);
        paragraph.appendText(" > ");
        paragraph.appendText("\"1\" ");
        paragraph.appendText("\"Greater than one\" ");
        paragraph.appendText("\"Less than one\"");

        FieldMark fieldMark = new FieldMark(document);
        fieldMark.setType(FieldMarkType.Field_End);

        paragraph.getItems().add(fieldMark);
        ifField.setEnd(fieldMark);
    }

    private static  void CreateIFField2(Document document, Paragraph paragraph)
    {

        IfField ifField = new IfField(document);
        ifField.setType(FieldType.Field_If);
        ifField.setCode("IF ");
        paragraph.getItems().add(ifField);

        paragraph.appendField("Age", FieldType.Field_Merge_Field);
        paragraph.appendText(" > ");
        paragraph.appendText("\"50\" ");
        paragraph.appendText("\"The old man\" ");
        paragraph.appendText("\"The young man\"");

        FieldMark fieldMark = new FieldMark(document);
        fieldMark.setType(FieldMarkType.Field_End);

        paragraph.getItems().add(fieldMark);
        ifField.setEnd(fieldMark);
    }
}
