import com.spire.doc.*;
import com.spire.doc.documents.*;
import com.spire.doc.fields.*;

public class createCrossReference {
    public static void main(String[] args) {
        //Create Word document.
        Document document = new Document();

        //Add a new section.
        Section section = document.addSection();

        //Create a bookmark.
        Paragraph paragraph = section.addParagraph();
        paragraph.appendBookmarkStart("MyBookmark");
        paragraph.appendText("Text inside a bookmark");
        paragraph.appendBookmarkEnd("MyBookmark");

        //Insert line breaks.
        for (int i = 0; i < 4; i++)
        {
            paragraph.appendBreak(BreakType.Line_Break);
        }

        //Create a cross-reference field, and link it to bookmark.
        Field field = new Field(document);
        field.setType(FieldType.Field_Ref);
        field.setCode("REF MyBookmark \\p \\h");

        //Insert field to paragraph.
        paragraph = section.addParagraph();
        paragraph.appendText("For more information, see ");
        paragraph.getChildObjects().add(field);

        //Insert FieldSeparator object.
        FieldMark fieldSeparator = new FieldMark(document, FieldMarkType.Field_Separator);
        paragraph.getChildObjects().add(fieldSeparator);

        //Set display text of the field.
        TextRange tr = new TextRange(document);
        tr.setText("above");
        paragraph.getChildObjects().add(tr);

        //Insert FieldEnd object to mark the end of the field.
        FieldMark fieldEnd = new FieldMark(document, FieldMarkType.Field_End);
        paragraph.getChildObjects().add(fieldEnd);

        String result = "output/CreateCrossReferenceToBookmark.docx";
        //Save to file.
        document.saveToFile(result, FileFormat.Docx_2013);
    }
}
