import com.spire.doc.*;
import com.spire.doc.documents.*;

public class moveToHeaderAndFooter {
    public static void main(String[] args) {
        // Create a new empty document instance.
        Document doc = new Document();

        // Create a document navigator to help navigate and modify the document content.
        DocumentNavigator navigator = new DocumentNavigator(doc);

        // Load an existing Word document from the specified relative file path.
        doc.loadFromFile("Data\\Sample.docx");

        // Move the navigator's cursor to the first section of the document (section index 0).
        navigator.moveToSection(0);

        // Navigate to the footer of the first page in the current section.
        navigator.moveToHeaderFooter(HeaderFooterType.Footer_First_Page);

        // Write a new line of text into the first-page footer.
        navigator.writeln("The footer on the first page.");

        // Navigate to the header of the first page in the current section.
        navigator.moveToHeaderFooter(HeaderFooterType.Header_First_Page);

        // Write a new line of text into the first-page header.
        navigator.writeln("The header on the first page.");

        // Save the modified document to a new file named "MoveToHeaderAndFooter.docx" in DOCX format.
        doc.saveToFile("MoveToHeaderAndFooter.docx", FileFormat.Docx);

        // Close the document to release internal resources.
        doc.close();

        // Explicitly dispose of the document object to free memory immediately.
        doc.dispose();
    }
}
