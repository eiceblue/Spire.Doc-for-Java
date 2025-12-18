import com.spire.doc.*;

public class ModifyBookmarkName {
    public static void main(String[] args) {

        // Create a new instance of the Document class
        Document document = new Document();

        // Load a Word document from the specified file path
        document.loadFromFile("Data/Bookmark.docx");

        // Retrieve the Bookmark object
        Bookmark bookmark = document.getBookmarks().get("Test");

        // Change the name of the retrieved bookmark to "bookmark1"
        bookmark.setName("bookmark1");

        // Save the modified document to a new file
        document.saveToFile("ModifyBookmarkName.docx", FileFormat.Docx_2019);

        // Close the document
        document.close();

        // Dispose of the document object
        document.dispose();
    }
}
