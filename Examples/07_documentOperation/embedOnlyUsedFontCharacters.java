import com.spire.doc.*;

public class embedOnlyUsedFontCharacters {
    public static void main(String[] args) {
        // Create a new instance of the Document class to handle the Word file.
        Document document = new Document();

        // Load the source Word document from the specified file path.
        document.loadFromFile("Data/Sample.docx");

        // Enable the option to embed fonts into the document file.
        document.setEmbedFontsInFile(true);

        // Configure the setting to embed only the characters used in the document (subset fonts) to reduce file size.
        document.setSaveSubsetFonts(true);

        // Save the modified document to a new file named "result.docx".
        document.saveToFile("result.docx", FileFormat.Docx_2019);

        // Close the document stream and release related resources.
        document.close();

        // Explicitly dispose of the document object to free up memory.
        document.dispose();
    }
}
