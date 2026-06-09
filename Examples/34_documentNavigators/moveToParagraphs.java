import com.spire.doc.*;
import com.spire.doc.documents.*;

public class moveToParagraphs {
    public static void main(String[] args) {
        // Create a new empty document instance.
        Document doc = new Document();

        // Create a document navigator to help navigate and modify the document content.
        DocumentNavigator navigator = new DocumentNavigator(doc);

        // Load an existing Word document from the specified relative file path.
        doc.loadFromFile("Data\\Sample.docx");

        // Move the navigator's cursor to the first section of the document (section index 0).
        navigator.moveToSection(0);

        // Move the cursor to the third paragraph (index 2) within the current section, at character offset 0.
        navigator.moveToParagraph(2, 0);

        // Insert new text at the current cursor position, overwriting any existing content from that point onward.
        navigator.writeln("This is new content......");

        // Save the modified document to a new file.
        doc.saveToFile("MoveToParagraphs.docx", FileFormat.Docx);

        // Close the document to release internal resources.
        doc.close();

        // Explicitly dispose of the document object to free memory immediately.
        doc.dispose();
    }
}
