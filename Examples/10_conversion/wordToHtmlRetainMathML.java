import com.spire.doc.*;

public class wordToHtmlRetainMathML {
    public static void main(String[] args) {
    
        // Initialize a new Document object
        Document document = new Document();

        // Load an existing Word document from the specified relative file path
        document.loadFromFile("Data/GetMathEquation.docx");

        // Retrieve the HTML export options configuration object for the document
        HtmlExportOptions htmlExportOptions = document.getHtmlExportOptions();

        // Configure the export to render Office math equations using MathML format
        htmlExportOptions.setOfficeMathOutputMode(HtmlOfficeMathOutputMode.Math_ML);

        // Set the CSS stylesheet to be embedded internally within the generated HTML file
        htmlExportOptions.setCssStyleSheetType(CssStyleSheetType.Internal);

        // Define the output file name for the converted HTML document
        String outputFile = "WordToHtmlRetainMathML.html";

        // Save the document as an HTML file using the configured export options
        document.saveToFile(outputFile, FileFormat.Html);

        // Close the document to release file handles
        document.close();

        // Dispose of the document object to free up memory
        document.dispose();
    }
}
