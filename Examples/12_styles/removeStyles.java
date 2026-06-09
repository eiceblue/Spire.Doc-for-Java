import com.spire.doc.*;

public class removeStyles {
    public static void main(String[] args) {
        // Initialize a new Document object to handle the Word file.
        Document document = new Document();

        // Load the Word document from the specified file path.
        document.loadFromFile("Data/RemoveStyles.docx");

        // Access the "Style1" style from the document's style collection and remove it completely.
        document.getStyles().get("Style1").removeSelf();

        // Access the "Style2" style from the document's style collection and remove it completely.
        document.getStyles().get("Style2").removeSelf();

        // Save the modified document to a new file in DOCX 2019 format.
        document.saveToFile("RemoveStyles-out.docx", FileFormat.Docx_2019);

        // Close the document to release file handles and resources associated with the open file.
        document.close();

        // Dispose of the document object to free up memory and clean up unmanaged resources.
        document.dispose();
    }
}
