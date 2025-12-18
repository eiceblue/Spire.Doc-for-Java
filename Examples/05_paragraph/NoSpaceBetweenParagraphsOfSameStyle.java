import com.spire.doc.*;
import com.spire.doc.documents.Paragraph;

public class NoSpaceBetweenParagraphsOfSameStyle {
    public static void main(String[] args) {

        // Create a new instance of the Document class
        Document document = new Document();

        // Load a Word document from the specified file path
        document.loadFromFile("Data/ExtractText.docx");

        // Get the Body object from the first section of the document
        Body body = document.getSections().get(0).getBody();

        // Loop through each paragraph in the body of the document
        Paragraph paragraph;
        for (int i = 0; i < body.getParagraphs().getCount(); i++) {
            // Retrieve the current paragraph
            paragraph = body.getParagraphs().get(i);

            // Set no space between paragraphs of the same style for the current paragraph
            paragraph.getFormat().setNoSpaceBetweenParagraphsOfSameStyle(true);
        }

        // Save the modified document to a new file with the specified format
        document.saveToFile("NoSpaceBetweenParagraphsOfSameStyle-out.docx", FileFormat.Docx_2019);

        // Dispose of the document object to release resources
        document.dispose();
    }
}
