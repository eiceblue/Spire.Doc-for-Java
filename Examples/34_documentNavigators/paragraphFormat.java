import com.spire.doc.*;
import com.spire.doc.documents.*;

import java.awt.*;

public class paragraphFormat {
    public static void main(String[] args) {
        // Create a new empty document instance.
        Document doc = new Document();

        // Create a document navigator to help navigate and modify the document content.
        DocumentNavigator navigator = new DocumentNavigator(doc);

        // Load an existing Word document from the specified relative file path.
        doc.loadFromFile("Data\\Sample.docx");

        // Move the navigator's cursor to the first section of the document (section index 0).
        navigator.moveToSection(0);

        // Move the cursor to the first paragraph (index 0) at character position 0 within that paragraph.
        navigator.moveToParagraph(0, 0);

        // Set the line spacing rule for the current paragraph to "Multiple" (enables custom line spacing multiplier).
        navigator.getParagraphFormat().setLineSpacingRule(LineSpacingRule.Multiple);

        // Set the line spacing to 1.5 times the default font size (assuming 12-point font: 1.5 * 12 = 18 points).
        navigator.getParagraphFormat().setLineSpacing(1.5F * 12F);

        // Set the left indent of the current paragraph to 5 points.
        navigator.getParagraphFormat().setLeftIndent(5);

        // Move the cursor to the third paragraph (index 2) at character position 0.
        navigator.moveToParagraph(2, 0);

        // Set the background color of the current paragraph to blue.
        navigator.getParagraphFormat().setBackColor(Color.BLUE);

        // Save the modified document to a new file.
        doc.saveToFile("ParagraphFormat.docx", FileFormat.Docx);

        // Close the document to release internal resources.
        doc.close();

        // Explicitly dispose of the document object to free memory immediately.
        doc.dispose();
    }
}
