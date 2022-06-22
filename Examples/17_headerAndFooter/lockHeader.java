import com.spire.doc.*;

public class lockHeader {
    public static void main(String[] args){
        String input = "data/headerSample.docx";
        String output = "output/lockHeader.docx";

        //load the document
        Document doc = new Document();
        doc.loadFromFile(input);

        //get the first section
        Section section = doc.getSections().get(0);

        //protect the document and set the ProtectionType as AllowOnlyFormFields
        doc.protect(ProtectionType.Allow_Only_Form_Fields, "123");

        //set the ProtectForm as false to unprotect the section
        section.setProtectForm(false);

        //save the document
        doc.saveToFile(output, FileFormat.Docx);
    }
}
