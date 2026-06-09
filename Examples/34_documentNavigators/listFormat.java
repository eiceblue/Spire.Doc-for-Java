import com.spire.doc.*;
import com.spire.doc.documents.*;

public class listFormat {
    public static void main(String[] args) {
        // Create a new document instance.
        Document doc = new Document();

        // Add a new section to the document.
        Section section = doc.addSection();

        // Add the first paragraph to the section.
        Paragraph paragraph = section.addParagraph();

        // Append text to the first paragraph.
        paragraph.appendText("This is the first paragraph.");

        // Add a second paragraph to the section.
        paragraph = section.addParagraph();

        // Append text to the second paragraph.
        paragraph.appendText("This is the second paragraph.");

        // Add a third paragraph to the section.
        paragraph = section.addParagraph();

        // Append text to the third paragraph.
        paragraph.appendText("This is the third paragraph.");

        // Create a document navigator to facilitate navigation and formatting.
        DocumentNavigator navigator = new DocumentNavigator(doc);

        // Apply bullet list style to the current position (first paragraph by default).
        navigator.getListFormat().applyBulletStyle();

        // Move the navigator's cursor to the third paragraph (index 2) at character offset 0.
        navigator.moveToParagraph(2, 0);

        // Apply bullet list style to the third paragraph.
        navigator.getListFormat().applyBulletStyle();

        // Save the document to a file named "ListFormat.docx" in DOCX format.
        doc.saveToFile("ListFormat.docx", FileFormat.Docx);

        // Close the document to release internal resources.
        doc.close();

        // Explicitly dispose of the document object to free memory immediately.
        doc.dispose();
    }
}
