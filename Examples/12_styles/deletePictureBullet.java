import com.spire.doc.*;
import com.spire.doc.collections.*;
import com.spire.doc.documents.ListLevel;

public class deletePictureBullet {
    public static void main(String[] args) {
        // Initialize a new Document object to work with Word documents.
        Document doc = new Document();

        // Load an existing Word document from the specified file path into the document object.
        doc.loadFromFile("Data/PictureBullets.docx");

        // Access the collection of list definitions (styles) within the document.
        ListCollection lists = doc.getListReferences();

        // Get the second level (index 1) of the first list definition (index 0) in the document.
        ListLevel listLevel = lists.get(0).getLevels().get(1);

        // Remove any picture bullet associated with this list level.
        listLevel.deletePictureBullet();

        // Save the modified document to a new file named "DeletePictureBullet.docx" in Word 2016 format.

        doc.saveToFile("DeletePictureBullet.docx", FileFormat.Docx_2016);
        // Close the document to release internal resources such as file handles.

        doc.close();

        // Explicitly dispose of the document object to ensure memory is freed immediately.
        doc.dispose();
    }
}
