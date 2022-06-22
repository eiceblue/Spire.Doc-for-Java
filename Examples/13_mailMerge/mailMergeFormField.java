import com.spire.doc.*;
import com.spire.doc.documents.Paragraph;
import com.spire.doc.fields.*;
import com.spire.doc.reporting.*;

import java.text.SimpleDateFormat;
import java.util.Date;

public class mailMergeFormField {

    public static void main(String[] args) throws Exception {

        String input = "data/MailMergeFormField.doc";
        String output = "output/MailMergeFormField_out.docx";

        Document doc = new Document();
        doc.loadFromFile(input);

        Date currentTime = new Date();
        SimpleDateFormat formatter = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
        String dateString = formatter.format(currentTime);

        String[] fieldNames = new String[] { "Contact Name", "Fax", "Date", "Urgent", "Share", "Submit","Body" };

        String[] fieldValues = new String[] { "John Smith", "+1 (69) 123456", dateString,
                "Yes","No","Yes",
                "<b>It's very urgent. Please deal with it ASAP. </b>" };
        doc.getMailMerge().MergeField = new MergeFieldEventHandler() {

            public void invoke(Object sender, MergeFieldEventArgs args) {
                mailMerge_MergeField(sender, args);
            }
        };

        //Mail merge
        doc.getMailMerge().execute(fieldNames, fieldValues);

        doc.saveToFile(output, FileFormat.Docx);

    }

    private static void mailMerge_MergeField(Object sender, MergeFieldEventArgs args) {
        if (args.getFieldValue().equals("Yes"))
        {
            //Create a checkbox name
            String checkBoxName = args.getFieldName();
            Paragraph para =  args.getCurrentMergeField().getOwnerParagraph();
            int index = para.getChildObjects().indexOf(args.getCurrentMergeField());
            // Insert a check box.
            CheckBoxFormField field = (CheckBoxFormField)para.appendField(checkBoxName, FieldType.Field_Form_Check_Box);
            para.getChildObjects().insert(index, field);
            para.getChildObjects().remove(args.getCurrentMergeField());
            field.setChecked(true);

        }
        if (args.getFieldValue().equals("No"))
        {
            //Create a checkbox name
            String checkBoxName = args.getFieldName();
            Paragraph para =  args.getCurrentMergeField().getOwnerParagraph();
            int index = para.getChildObjects().indexOf(args.getCurrentMergeField());
            // Insert a check box.
            CheckBoxFormField field = (CheckBoxFormField)para.appendField(checkBoxName, FieldType.Field_Form_Check_Box);
            para.getChildObjects().insert(index, field);
            para.getChildObjects().remove(args.getCurrentMergeField());
            field.setChecked(false);
        }
        // Insert html during mail merge.
        if (args.getFieldName().equals("Body"))
        {
            Paragraph para =  args.getCurrentMergeField().getOwnerParagraph();
            para.appendHTML(args.getFieldValue().toString());
            para.getChildObjects().remove(args.getCurrentMergeField());
        }

        // Insert text input form field.
        if (args.getFieldName().equals("Date"))
        {
            String textInputName = args.getFieldName();
            Paragraph para =  args.getCurrentMergeField().getOwnerParagraph();
            TextFormField field = (TextFormField)para.appendField(textInputName, FieldType.Field_Form_Text_Input);
            para.getChildObjects().remove(args.getCurrentMergeField());
            field.setText(args.getFieldValue().toString());
        }
    }
}
