import com.spire.doc.*;
import com.spire.doc.documents.comparison.*;

import java.util.Date;

public class documentComparisonEnhanced {
    public static void main(String[] args) {
        // Create a new Document object for the first file
        Document doc1 = new Document();
        // Load the first document from the specified file path
        doc1.loadFromFile("Data/DocumentComparisonEnhanced1.docx");

        // Create a new Document object for the second file
        Document doc2 = new Document();
        // Load the second document from the specified file path
        doc2.loadFromFile("Data/DocumentComparisonEnhanced2.docx");

        // Initialize comparison options to customize the comparison behavior
        CompareOptions options = new CompareOptions();
        // Enable tracking of moved content during comparison
        options.setCompareMoves(true);
        // Ignore changes in letter casing (e.g., 'Word' vs 'word')
        options.setIgnoreCaseChanges(true);
        // Ignore any differences in comments
        options.setIgnoreComments(true);
        // Ignore changes within document fields
        options.setIgnoreFields(true);
        // Ignore differences in footnotes and endnotes
        options.setIgnoreFootnotes(true);
        // Ignore changes within tables
        options.setIgnoreTables(true);
        // Ignore differences in text boxes
        options.setIgnoreTextboxes(true);

        // Compare the two documents, marking changes with the author name "administrator" and current date
        doc1.compare(doc2, "administrator", new Date(), options);
        // Save the comparison result to a new file in Word 2013 format
        doc1.saveToFile("DocumentComparisonEnhanced.docx", FileFormat.Docx_2013);

        // Close both documents to release file handles
        doc1.close();
        doc2.close();

        // Dispose of both document objects to free up memory resources
        doc1.dispose();
        doc2.dispose();
    }
}
