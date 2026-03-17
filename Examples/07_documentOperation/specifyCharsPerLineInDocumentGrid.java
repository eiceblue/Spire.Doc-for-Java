import com.spire.doc.*;
import com.spire.doc.documents.*;

public class specifyCharsPerLineInDocumentGrid {
    public static void main(String[] args) {
        // Create a new document instance
        Document document = new Document();

        // Add a new section to the document
        Section section = document.addSection();

        // Set the document grid type to character and line grid
        section.getPageSetup().setGridType(GridPitchType.Chars_And_Line);

        // Specify the number of characters per line in the document grid
        section.getPageSetup().setCharactersPerLine(30);

        // Add a new paragraph to the section
        Paragraph paragraph = section.addParagraph();

        // Append sample text to the paragraph
        paragraph.appendText("Spire.Doc for Java is a professional Word API that empowers Java applications to create, convert, manipulate and print Word documents without dependency on Microsoft Word.");

        // Save the document to a .docx file
        document.saveToFile("SpecifyCharsPerLineInDocumentGrid.docx", FileFormat.Docx);

        // Close the document to release resources
        document.close();

        // Dispose of the document object to free memory
        document.dispose();
    }
}
