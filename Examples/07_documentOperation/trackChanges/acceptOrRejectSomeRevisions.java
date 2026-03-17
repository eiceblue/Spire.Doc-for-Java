import com.spire.doc.*;

public class acceptOrRejectSomeRevisions {
    public static void main(String[] args) {
        // Create a new Document instance to work with a Word document in memory.
        Document document = new Document();

        // Load an existing Word document.
        document.loadFromFile("Data/GetRevisions.docx");

        // Retrieve the collection of all revision information (track changes) present in the document.
        RevisionInfoCollection revisionInfoCollection = document.getRevisionInfos();

        // Iterate through each revision in the document using a for loop.
        for (int i = 0; i < revisionInfoCollection.getCount(); i++) {
            // Get the revision info object at the current index.
            RevisionInfo revisionInfo = revisionInfoCollection.get(i);

            // Check if the current revision is an insertion (newly added content).
            if (revisionInfo.getRevisionType() == RevisionType.Insertion) {

                // Accept (keep) the insertion revision in the document.
                revisionInfo.accept();

                // Alternative: reject (remove) the insertion instead of accepting it.
                // revisionInfo.reject();

                // Decrement the loop counter to re-check the current index,
                // because accepting/rejecting a revision may alter the revision collection.
                i--;
            }
        }

        // Save the modified document (with accepted/rejected revisions) to a new file.
        document.saveToFile("AcceptOrRejectSomeRevisions.docx", FileFormat.Docx);

        // Close the document to release internal resources.
        document.close();

        // Explicitly dispose of the document object to free memory and system handles.
        document.dispose();
    }
}
