import com.spire.doc.*;
import com.spire.doc.documents.*;

import java.awt.*;

public class characterformat {
    public static void main(String[] args) {
        // Load the document by creating a new Document instance.
        Document doc = new Document();

        // Create a DocumentNavigator object to facilitate easy content insertion and formatting.
        DocumentNavigator navigator = new DocumentNavigator(doc);

        // Write plain text into the document (without special formatting yet).
        navigator.writeln("Write plain text into the document (without special formatting yet).");

        // Enable underline formatting for subsequent text.
        navigator.getCharacterFormat().setUnderlineColor(SystemColor.ORANGE);
        navigator.getCharacterFormat().setUnderlineStyle(UnderlineStyle.Single);

        // Set bold formatting for subsequent text.
        navigator.getCharacterFormat().setBold(true);

        // Enable shadow effect for subsequent text.
        navigator.getCharacterFormat().isShadow(true);

        // Set the text color to blue for subsequent text.
        navigator.getCharacterFormat().setTextColor(Color.BLUE);

        // Write formatted text using the current character formatting settings.
        navigator.writeln("Write formatted text using the current character formatting settings.");

        // Save the current character formatting settings onto an internal stack for later reuse.
        navigator.pushCharacterFormat();

        // Clear all character formatting to default (e.g., no bold, no color, etc.).
        navigator.getCharacterFormat().clearFormatting();

        // Write text with cleared (default) formatting.
        navigator.writeln("Write text with cleared (default) formatting");

        // Restore the previously saved character formatting from the stack.
        navigator.popCharacterFormat();

        // Write text using the restored formatting.
        navigator.writeln("Write text using the restored formatting.");

        // Save the document to a file in DOCX format.
        doc.saveToFile("Characterformat.docx", FileFormat.Docx);

        // Close the document to release internal resources.
        doc.close();

        // Explicitly dispose of the document object to free memory immediately.
        doc.dispose();
    }
}
