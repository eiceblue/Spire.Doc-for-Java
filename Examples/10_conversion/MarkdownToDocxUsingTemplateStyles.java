import com.spire.doc.*;

public class MarkdownToDocxUsingTemplateStyles {
    public static void main(String[] args) {

        // Initialize a new Document object by loading the Markdown file from the specified relative path.
        Document doc = new Document("Data\\sample.md");

        // Copy all styles from the specified Word template into the current document.
        doc.copyStylesFromTemplate("Data\\template.docx");

        // Define the output filename for the converted Word document.
        String outputFile = "MarkdownToDocxUsingTemplateStyles.docx";

        // Save the processed document to the specified file in DOCX 2016 format.
        doc.saveToFile(outputFile, FileFormat.Docx_2016);

        // Close the document to release resources.
        doc.close();
    }
}
