import com.spire.doc.*;

public class specifiedProtectionType {
    public static void main(String[] args) {
        //Create Word document.
        Document document = new Document();

        //Load the file from disk.
        document.loadFromFile("data/JAVATemplate_N.docx");

        //Protect the Word file.
        document.protect(ProtectionType.Allow_Only_Reading, "123456");

        //Save to file.
        String result = "output/SpecifiedProtectionType.docx";
        document.saveToFile(result, FileFormat.Docx_2013);
    }
}
