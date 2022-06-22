import com.spire.doc.*;

public class lockSpecifiedSections {
    public static void main(String[] args) {
        //Create Word document.
        Document document = new Document();

        //Add new sections.
        Section s1 = document.addSection();
        Section s2 = document.addSection();

        //Append some text to section 1 and section 2.
        s1.addParagraph().appendText("Spire.Doc demo, section 1");
        s2.addParagraph().appendText("Spire.Doc demo, section 2");

        //Protect the document with AllowOnlyFormFields protection type.
        document.protect(ProtectionType.Allow_Only_Form_Fields, "123");

        //Unprotect section 2
        s2.setProtectForm(false);

        //Save to file.
        String result = "output/LockSpecifiedSections.docx";
        document.saveToFile(result, FileFormat.Docx_2013);
    }
}
