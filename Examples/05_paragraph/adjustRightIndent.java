import com.spire.doc.*;
import com.spire.doc.documents.*;

public class adjustRightIndent {
    public static void main(String[] args) {
    
        // Create a new instance of the Document class
        Document doc = new Document();

        // Add a new section to the document
        Section section = doc.addSection();

        // Add a new paragraph to the body of the section
        Paragraph paragraph = section.getBody().addParagraph();

        // Set the text content of the paragraph
        paragraph.setText("Hello World!");

        // Enable the adjustment of the right indent for the paragraph format
        paragraph.getFormat().setAdjustRightIndent(true);

        // Add another new paragraph to the body of the section
        paragraph = section.getBody().addParagraph();

        // Set the text content for the second paragraph
        paragraph.setText("Thank you for using the Spire.Doc product.");

        // Disable the adjustment of the right indent for this paragraph
        paragraph.getFormat().setAdjustRightIndent(false);

        // Define the file path and name for the output document
        String result = "AdjustRightIndent.docx";

        // Save the document to a file in Docx 2016 format
        doc.saveToFile(result, FileFormat.Docx_2016);

        // Close the document to release resources
        doc.close();

        // Dispose of the document object to free up memory
        doc.dispose();
    }
}
