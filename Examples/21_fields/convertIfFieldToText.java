import com.spire.doc.*;
import com.spire.doc.collections.FieldCollection;
import com.spire.doc.documents.*;
import com.spire.doc.fields.*;

public class convertIfFieldToText {
    public static void main(String[] args) {
        //Open a Word document
        Document document = new Document("data/IfFieldSample.docx");

        //Get all fields in document
        FieldCollection fields = document.getFields();

        for (int i = 0; i < fields.getCount(); i++)
        {
            Field field = fields.get(i);
            if (field.getType().equals(FieldType.Field_If))
            {
                TextRange original = (TextRange)field;
                //Get field text
                String text = field.getFieldText();
                //Create a new textRange and set its format
                TextRange textRange = new TextRange(document);
                textRange.setText(text);
                textRange.getCharacterFormat().setFontName(original.getCharacterFormat().getFontName());
                textRange.getCharacterFormat().setFontSize(original.getCharacterFormat().getFontSize());

                Paragraph par = field.getOwnerParagraph();
                //Get the index of the if field
                int index = par.getChildObjects().indexOf(field);
                //Remove if field via index
                par.getChildObjects().removeAt(index);
                //Insert field text at the position of if field
                par.getChildObjects().insert(index, textRange);
            }
        }

        String result ="output/ConvertIfFieldToText.docx";
        //Save doc file
        document.saveToFile(result, FileFormat.Docx);
    }
}
