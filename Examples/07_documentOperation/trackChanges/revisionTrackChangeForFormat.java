import com.spire.doc.*;
import com.spire.doc.documents.DocumentObjectType;
import com.spire.doc.documents.Paragraph;
import com.spire.doc.fields.TextRange;
import java.awt.*;

public class revisionTrackChangeForFormat {
    public static void main(String[] args) {
        // Create a new Document instance.
        Document document = new Document();

        // Load an existing Word document.
        document.loadFromFile("Data/Template.docx");

        // Enable revision tracking (Track Changes).
        document.startTrackRevisions("SpireDoc");

        // Get the first section of the document.
        Section section = document.getSections().get(0);

        // Get the first paragraph in that section.
        Paragraph paragraph = section.getParagraphs().get(0);

        // Loop through all child objects (e.g., text ranges, images, etc.) in the first paragraph.
        for (int i = 0; i < paragraph.getChildObjects().getCount(); i++) {
            // Check if the current child object is a text range (i.e., actual textual content).
            if (paragraph.getChildObjects().get(i).getDocumentObjectType() == DocumentObjectType.Text_Range) {
                // Cast the object to a TextRange to access its formatting properties.
                TextRange tr = (TextRange) paragraph.getChildObjects().get(i);

                // Change the text color to red—this will be recorded as a formatting revision.
                tr.getCharacterFormat().setTextColor(Color.RED);

                // Set the font size to 28 points—tracked as a revision due to active track changes.
                tr.getCharacterFormat().setFontSize(28);

                // Apply bold formatting—also captured as a tracked change.
                tr.getCharacterFormat().setBold(true);
            }
        }

        // Append new text to the second paragraph —this insertion will be tracked as a revision.
        section.getParagraphs().get(1).appendText("some new text");

        // Disable revision tracking (stop recording changes).
        document.stopTrackRevisions();

        // Save the modified document with tracked revisions to a new file in .docx format.
        document.saveToFile("revisionTrackChangeForFormat-out.docx", FileFormat.Docx);

        // Close the document to release internal resources.
        document.close();

        // Explicitly dispose of the document object to free memory and system handles.
        document.dispose();
    }
}
