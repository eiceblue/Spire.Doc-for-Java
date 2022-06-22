import com.spire.doc.*;
public class loadAndSaveToDisk {
    public static void main(String[] args) {
        String input = "data/Template.docx";
        //Create a new document
        Document doc = new Document();
        // Load the document from the absolute/relative path on disk.
        doc.loadFromFile(input);

        String result = "output/loadAndSaveToDisk_out.docx";
        // Save the document to disk
        doc.saveToFile(result,FileFormat.Docx);
    }
}
