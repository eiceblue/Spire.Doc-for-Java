import com.spire.doc.*;
import com.spire.doc.documents.*;

public class insertTextFormField {
    public static void main(String[] args) {
        // Create a new empty Word document.
        Document doc = new Document();

        // Initialize a DocumentNavigator to help insert content and form fields into the document.
        DocumentNavigator navigator = new DocumentNavigator(doc);

        // Write a label indicating the type of the following form field: Calculation.
        navigator.writeln("TextFormFieldType.Calculation: ");

        // Insert a calculation-type text form field.
        navigator.insertTextFormField("CalculationTextField", TextFormFieldType.Calculation, "0", "=3+1", 30);

        // Insert a line break (carriage return) to move to the next line.
        navigator.writeln();

        // Write a label for a number-only text form field.
        navigator.writeln("TextFormFieldType.NumberText: ");

        // Insert a number-input text form field.
        navigator.insertTextFormField("NumberText", TextFormFieldType.Number_Text, "0", "100", 30);

        // Insert a line break.
        navigator.writeln();

        // Write a label for a date-input text form field.
        navigator.writeln("TextFormFieldType.DateText: ");

        // Insert a date-formatted text form field.
        navigator.insertTextFormField("DateText", TextFormFieldType.Date_Text, "yyyy/M/d", "2025/8/1", 30);

        // Insert a line break.
        navigator.writeln();

        // Enable automatic field updating.
        doc.isUpdateFields(true);

        // Save the resulting document to a file.
        doc.saveToFile("InsertTextFormField.docx", FileFormat.Docx);

        // Close the document to release internal resources.
        doc.close();

        // Explicitly dispose of the document object to free memory immediately.
        doc.dispose();
    }
}
