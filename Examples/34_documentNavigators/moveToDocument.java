import com.spire.doc.Document;
import com.spire.doc.FileFormat;
import com.spire.doc.documents.DocumentNavigator;

public class moveToDocument {
    public static void main(String[] args) {
        // Create a new empty document instance.
        Document doc = new Document();

        // Create a document navigator to help navigate and modify the document content.
        DocumentNavigator navigator = new DocumentNavigator(doc);

        // Load an existing Word document from the specified relative file path.
        doc.loadFromFile("Data//Sample.docx");

        // Move the cursor to the very beginning of the document.
        navigator.moveToDocumentStart();

        // Write a new line of text at the start of the document.
        navigator.writeln("Insert the content at the beginning of the document.");

        // Write another line of text immediately after the previous one at the start.
        navigator.writeln("This is new content.");

        // Move the cursor to the very end of the document.
        navigator.moveToDocumentEnd();

        // Insert a blank line at the end of the document.
        navigator.writeln();

        // Insert a new line of text at the end of the document.
        navigator.writeln("Insert the content at the end of the document.");

        // Save the modified document to a new file.
        doc.saveToFile("MoveToDocument.docx", FileFormat.Docx);

        // Close the document to release internal resources.
        doc.close();

        // Explicitly dispose of the document object to free memory immediately.
        doc.dispose();
    }
}
