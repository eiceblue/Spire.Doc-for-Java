import com.spire.doc.*;
import com.spire.doc.documents.*;

public class pageFormat {
    public static void main(String[] args) {
        // Create a new empty document instance.
        Document doc = new Document();

        // Create a document navigator to help navigate and modify the document content.
        DocumentNavigator navigator = new DocumentNavigator(doc);

        // Load an existing Word document from the specified relative file path.
        doc.loadFromFile("Data\\Sample.docx");

        // Move the navigator's cursor to the first section (section index 0) of the document.
        navigator.moveToSection(0);

        // Set the page margins for the current section.
        navigator.getPageSetup().setMargins(new MarginsF(100, 80, 100, 80));

        // Set the page size of the current section to Letter.
        navigator.getPageSetup().setPageSize(PageSize.Letter);

        // Save the modified document to a new file.
        doc.saveToFile("PageFormat.docx", FileFormat.Docx);

        // Close the document to release internal resources.
        doc.close();

        // Explicitly dispose of the document object to free memory immediately.
        doc.dispose();

    }
}
