import com.spire.doc.*;
import com.spire.doc.fields.FormField;

import java.awt.*;

public class formFieldsProperties {
    public static void main(String[] args) {
        //Open a Word document
        Document document = new Document("data/FillFormField.doc");

        //Get the first section
        Section section = document.getSections().get(0);

        //Get FormField by index
        FormField formField = section.getBody().getFormFields().get(1);

        if (formField.getType().equals(FieldType.Field_Form_Text_Input))
        {
            formField.setText("My name is " + formField.getName());
            formField.getCharacterFormat().setTextColor(Color.red);
            formField.getCharacterFormat().setItalic(true);
        }

        //save to file
        document.saveToFile("output/formFieldsProperties.docx", FileFormat.Docx);
    }
}
