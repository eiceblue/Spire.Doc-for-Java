import com.spire.doc.*;
import com.spire.doc.fields.TextRange;

public class obtainFormatChangeRevisions {
    public static void main(String[] args) {
        // Create a new Document instance to work with a Word document in memory.
        Document document = new Document();

        // Load an existing Word document that contains tracked changes (revisions) from the specified file path.
        document.loadFromFile("Data/GetRevisions.docx");

        // Retrieve the collection of all revision information (i.e., tracked changes such as insertions, deletions, or formatting changes) in the document.
        RevisionInfoCollection revisionInfoCollection = document.getRevisionInfos();

        // Iterate through each revision in the collection using a standard for loop.
        for (int i = 0; i < revisionInfoCollection.getCount(); i++) {
            // Obtain the specific revision info object at the current index.
            RevisionInfo revisionInfo = revisionInfoCollection.get(i);

            // Check if the current revision represents a formatting change (e.g., font size, color, style modifications).
            if (revisionInfo.getRevisionType() == RevisionType.Format_Change) {

                // Verify that the object affected by the formatting change is a text range (i.e., actual textual content).
                if (revisionInfo.getOwnerObject() instanceof TextRange) {
                    // Cast the owner object to a TextRange to access its text and formatting properties.
                    TextRange textRange = (TextRange) revisionInfo.getOwnerObject();

                    // Print the text content of the formatted segment.
                    System.out.println(textRange.getText());

                    // Switch the document's revision view mode to show the original version (before the formatting change).
                    document.setRevisionsView(RevisionsView.Original);

                    // Print the font size as it appeared in the original version.
                    System.out.println(textRange.getCharacterFormat().getFontSize());

                    // Switch the document's revision view mode to show the final version (with the formatting change applied).
                    document.setRevisionsView(RevisionsView.Final);

                    // Print the font size as it appears after the formatting revision.
                    System.out.println(textRange.getCharacterFormat().getFontSize());
                }
            }
        }
    }
}