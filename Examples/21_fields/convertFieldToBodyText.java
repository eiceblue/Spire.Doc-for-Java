import com.spire.doc.*;
import com.spire.doc.documents.*;
import com.spire.doc.fields.*;

public class convertFieldToBodyText {
    public static void main(String[] args) {
        //Create the source document
        Document sourceDocument = new Document();

        //Load the source document from disk.
        sourceDocument.loadFromFile("data/TextInputField.docx");

        //Traverse FormFields
        for (FormField field : (Iterable<FormField>)sourceDocument.getSections().get(0).getBody().getFormFields())
        {
            //Find FieldFormTextInput type field
            if (field.getType().equals(FieldType.Field_Form_Text_Input))
            {
                //Get the paragraph
                Paragraph paragraph = field.getOwnerParagraph();

                //Define variables
                int startIndex = 0;
                int endIndex = 0;

                //Create a new TextRange
                TextRange textRange = new TextRange(sourceDocument);

                //Set text for textRange
                textRange.setText(paragraph.getText());

                //Traverse DocumentObjectS of field paragraph
                for(DocumentObject obj :(Iterable<DocumentObject>)paragraph.getChildObjects())
                {
                    //If its DocumentObjectType is BookmarkStart
                    if (obj.getDocumentObjectType().equals(DocumentObjectType.Bookmark_Start))
                    {
                        //Get the index
                        startIndex = paragraph.getChildObjects().indexOf(obj);
                    }
                    //If its DocumentObjectType is BookmarkEnd
                    if (obj.getDocumentObjectType().equals(DocumentObjectType.Bookmark_End))
                    {
                        //Get the index
                        endIndex = paragraph.getChildObjects().indexOf(obj);
                    }
                }
                //Remove ChildObjects
                for (int i = endIndex; i > startIndex; i--)
                {
                    //If it is TextFormField
                    if (paragraph.getChildObjects().get(i) instanceof TextFormField)
                    {
                        TextFormField textFormField = (TextFormField)paragraph.getChildObjects().get(i);

                        //Remove the field object
                        paragraph.getChildObjects().remove(textFormField);
                    }
                    else
                    {
                        paragraph.getChildObjects().removeAt(i);
                    }
                }
                //Insert the new TextRange
                paragraph.getChildObjects().insert(startIndex, textRange);
                break;
            }

        }
        //Save the document.
        String output="output/ConvertFieldToBodyText.docx";
        sourceDocument.saveToFile(output, FileFormat.Docx);
    }
}
