import com.spire.doc.*;
import java.util.Date;

public class StartTrackRevisions {
    public static void main(String[] args) {
        // Create a new Document object
        Document document=new Document();

        // Load the document from the specified file path
        document.loadFromFile("Data/ExtractText.docx");

        // Start the track revisions
        document.startTrackRevisions("User01", new Date());

        // Get the first paragraph and add content
        document.getSections().get(0).getParagraphs().get(0).appendText ("User01 add new Text!");

        // Delete a paragraph
        document.getSections().get(0).getParagraphs().removeAt(2);

        // Stop the track revisions
        document.stopTrackRevisions();

        // Save the file
        document.saveToFile("StartTrackRevisions_out.docx", FileFormat.Docx);

        // Dispose of the Document object
        document.dispose();
    }
}
