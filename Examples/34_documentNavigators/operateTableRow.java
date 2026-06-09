import com.spire.doc.*;
import com.spire.doc.*;
import com.spire.doc.documents.*;

import java.awt.*;

public class operateTableRow {
    public static void main(String[] args) {
        // Create a new empty document instance.
        Document doc = new Document();

        // Create a document navigator to help navigate and modify the document content.
        DocumentNavigator navigator = new DocumentNavigator(doc);

        // Start creating a new table and get its reference.
        Table table = navigator.startTable();

        // Initialize the table with 2 rows and 2 columns.
        table.resetCells(2, 2);

        // Set the table width to 100% of the available page width.
        table.setPreferredWidth(new PreferredWidth(WidthType.Percentage, (short)100));

        // Set the height of the first row to 30 points.
        table.getFirstRow().setHeight(30f);

        // Get a reference to the first cell in the first row (row 0, column 0).
        TableCell cell1 = table.getRows().get(0).getCells().get(0);

        // Add content to the first cell of the first row.
        addcellcontent(cell1, "Row 1, Cell 1");

        // Get a reference to the second cell in the first row (row 0, column 1).
        TableCell cell2 = table.getRows().get(0).getCells().get(1);

        // Add content to the second cell of the first row.
        addcellcontent(cell2, "Row 1, Cell 2");

        // Get a reference to the first cell in the second row (row 1, column 0).
        TableCell cell3 = table.getRows().get(1).getCells().get(0);

        // Add content to the first cell of the second row.
        addcellcontent(cell3, "Row 2, Cell 1");

        // Get a reference to the second cell in the second row (row 1, column 1).
        TableCell cell4 = table.getRows().get(1).getCells().get(1);

        // Add content to the second cell of the second row.
        addcellcontent(cell4, "Row 2, Cell 2");

        // Finalize the table creation.
        navigator.endTable();

        // Delete the first row (row index 0) of the first table (table index 0) in the document.
        navigator.deleteRow(0, 0);

        // Save the modified document to a new file named "OperateTableRow.docx" in DOCX format.
        doc.saveToFile("OperateTableRow.docx", FileFormat.Docx);

        // Close the document to release internal resources.
        doc.close();

        // Explicitly dispose of the document object to free memory immediately.
        doc.dispose();
    }
    static void addcellcontent(TableCell cell, String Content)
    {
        // Add a new paragraph to the table cell.
        Paragraph para = cell.addParagraph();

        // Append the specified text content to the paragraph inside the cell.
        para.appendText(Content);

        // Center-align the text horizontally within the paragraph.
        para.getFormat().setHorizontalAlignment(HorizontalAlignment.Center);

        // Set the background (shading) color of the table cell to ORANGE.
        cell.getCellFormat().getShading().setBackgroundPatternColor(Color.ORANGE);

        // Vertically center the content within the table cell.
        cell.getCellFormat().setVerticalAlignment(VerticalAlignment.Middle);
    }
}
