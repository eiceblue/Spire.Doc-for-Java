import com.spire.doc.*;

public class odtToWord {
    public static void main(String[] args) {
        //Create Word document.
        Document document = new Document();

        //Load the file from disk.
        document.loadFromFile("data/Template_OdtFile.odt");

        String result = "output/result-OdtToDocOrDocx.docx";

        //Save to Docx file.
        document.saveToFile(result, FileFormat.Docx);
    }
}
