import com.spire.doc.Document;

/**
 * Check the protection password for Word file.
 */
public class checkProtectionPassword {
    public static void main(String []args){
        String input="data/checkProtectionPassword.docx";
        String password="e-iceblue";
        //Create a new document
        Document document = new Document();
        document.loadFromFile(input);
        //Check the protection password for the document
        boolean checkResult = document.checkProtectionPassWord(password);
    }
}
