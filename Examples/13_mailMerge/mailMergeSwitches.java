import com.spire.doc.*;

public class mailMergeSwitches {
    public static void main(String[] args) throws Exception {
        String input = "data/MailMergeSwitches.docx";
        String output = "output/mailMergeSwitches result.docx";

        Document document = new Document();
        document.loadFromFile(input);


        String[] fieldName = new String[] { "XX_Name" };
        String[] fieldValue = new String[] { "Jason Tang" };

        document.getMailMerge().execute(fieldName, fieldValue);
        
        document.saveToFile(output, FileFormat.Docx);
    }
}
