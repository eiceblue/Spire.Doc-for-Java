import com.spire.doc.*;
import com.spire.doc.documents.*;

public class insertField {
    public static void main(String[] args) {
        // Create a new DocumentNavigator instance, which automatically creates an underlying empty document.
        DocumentNavigator navigator = new DocumentNavigator();

        // Get the Document object associated with the navigator.
        Document doc = navigator.getDocument();

        // Write the label text "Add page fields:" followed by a paragraph break.
        navigator.writeln("Add page fields:");

        // Insert a PAGE field with numeric formatting.
        navigator.insertField("PAGE \\# \"Page 0\"");

        // Insert a paragraph break (empty line).
        navigator.writeln();

        // Insert the same PAGE field but with a custom result text to simulate a placeholder or initial value.
        navigator.insertField("PAGE \\# \"Page 0\"", "3");

        // Insert another paragraph break.
        navigator.writeln();

        // Insert a built-in PAGE field using FieldType enumeration, with the field result displayed (true = show result).
        navigator.insertField(FieldType.Field_Page, true);

        // Insert another paragraph break.
        navigator.writeln();

        // Insert another PAGE field, but this time hide the field result (false = show field code instead of result).
        navigator.insertField(FieldType.Field_Page, false);

        // Save the document to a file in DOCX format.
        doc.saveToFile("InsertField.docx", FileFormat.Docx);

        // Close the document to release internal resources.
        doc.close();

        // Explicitly dispose of the document object to free memory immediately.
        doc.dispose();
    }
}
