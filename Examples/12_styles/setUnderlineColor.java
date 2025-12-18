import com.spire.doc.*;
import com.spire.doc.documents.*;
import com.spire.doc.fields.*;
import java.awt.*;

public class setUnderlineColor {
    public static void main(String[] args) {
        // Create a new Document instance
        Document document = new Document();

        // Add a new section to the document
        Section section = document.addSection();

        // Add a new paragraph to the section
        Paragraph paragraph = section.addParagraph();

        // Append text to the paragraph and get the TextRange object for formatting
        TextRange textRange = paragraph.appendText("Welcome to evaluate Spire.Doc for Java product.");

        // Set the underline style of the text to single underline
        textRange.getCharacterFormat().setUnderlineStyle(UnderlineStyle.Single);

        // Set the underline color of the text to red
        textRange.getCharacterFormat().setUnderlineColor(Color.red);

        // Define the file path and name for saving the document
        String filePath = "output/SetColorOfUnderline.docx";

        // Save the document to the specified file path in DOCX format
        document.saveToFile(filePath, FileFormat.Docx);

        // Release resources used by the document
        document.dispose();
    }
}
