import com.spire.doc.*;
import com.spire.doc.documents.comparison.CompareOptions;

public class compareDocumentsIgnoreTable {
    public static void main(String[] args) {
        // Load the first document from the specified file path
        Document document1 = new Document("data/ComparedDoc1.docx");

        // Load the second document from the specified file path
        Document document2 = new Document("data/ComparedDoc2.docx");

        // Create a new CompareOptions object to specify comparison settings
        CompareOptions compareoptions = new CompareOptions();

        // Set the option to ignore differences in tables during comparison
        compareoptions.setIgnoreTable(true);

        // Compare the two documents using the specified options, with "E-iceblue" as the author name for tracked changes
        document1.compare(document2, "E-iceblue", compareoptions);

        // Save the compared document (with changes tracked) to a new file in DOCX 2019 format
        document1.saveToFile("output/CompareDocumentsIgnoreTable.docx", FileFormat.Docx_2019);

        // Release resources used by the first document
        document1.dispose();

        // Release resources used by the second document
        document2.dispose();
    }
}
