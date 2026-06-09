import com.spire.doc.*;
import com.spire.doc.documents.*;
import java.awt.*;

public class applyCustomStylesToTable {
    public static void main(String[] args) {
        // Initialize a new Document object to create a fresh Word file.
        Document doc = new Document();

        // Add a new section to the document, which is required before adding content like tables.
        Section section = doc.addSection();

        // Create a new custom table style named "TestTableStyle1" and cast it to TableStyle type.
        TableStyle tableStyle = (TableStyle) doc.getStyles().add(StyleType.Table_Style, "TestTableStyle1");

        // Set the horizontal alignment of the table content to Center.
        tableStyle.setHorizontalAlignment(RowAlignment.Center);

        // Access the borders of the style and set their color to Blue.
        tableStyle.getBorders().setColor(Color.BLUE);

        // Set the border line style to Single (a solid thin line).
        tableStyle.getBorders().setBorderType(BorderStyle.Single);

        // Add a new table to the current section.
        Table table = section.addTable();

        // Reset the table structure to have 1 row and 1 column.
        table.resetCells(1, 1);

        // Access the first cell (Row 0, Cell 0), add a paragraph to it, and append the specific text.
        table.getRows().get(0).getCells().get(0).addParagraph().appendText("Aligned to the center of the page");

        // Set the preferred width of the table to 300 points.
        table.setPreferredWidth(PreferredWidth.fromPoints(300));

        // Apply the previously defined custom style ("TestTableStyle1") to this table.
        table.applyStyle(tableStyle);

        // Save the document to a file named "applyCustomStylesToTable.docx" in standard DOCX format.
        doc.saveToFile("applyCustomStylesToTable.docx", FileFormat.Docx);

        // Close the document to release file handles.
        doc.close();

        // Dispose of the document object to free up memory resources.
        doc.dispose();
    }
}
