import com.spire.doc.*;
import com.spire.doc.documents.*;
import com.spire.doc.formatting.CellFormat;

import java.awt.*;

public class operateTableCell {
    public static void main(String[] args) {
        // Create a new document instance.
        Document doc = new Document();

        // Initialize a document navigator to navigate and modify the document content.
        DocumentNavigator navigator = new DocumentNavigator(doc);

        // Begin creating a new table in the document.
        navigator.startTable();

        // Insert the first cell into the current row of the table and get its reference.
        TableCell cell1 = navigator.insertCell();

        // Add content to the first cell of the first row.
        addcellcontent(cell1, "Row 1, Cell 1");

        // Insert the second cell into the current row of the table and get its reference.
        TableCell cell2 = navigator.insertCell();

        // Add content to the second cell of the first row.
        addcellcontent(cell2, "Row 1, Cell 2");

        // Insert the third cell into the current row of the table and get its reference.
        TableCell cell3 = navigator.insertCell();

        // Add content to the third cell of the first row.
        addcellcontent(cell3, "Row 1, Cell 3");

        // End the current row and move to the next row in the table.
        navigator.endRow();

        // Insert the first cell into the new row of the table and get its reference.
        TableCell cell4 = navigator.insertCell();

        // Add content to the first cell of the second row.
        addcellcontent(cell4, "Row 2, Cell 1");

        // End the table creation process.
        navigator.endTable();

        // Move the navigator's cursor to a specific cell in the first table
        navigator.moveToCell(0, 0, 1, 0);

        // Insert (overwrite) the text at the current cursor position inside the target cell.
        navigator.writeln("new content");

        // Get the formatting object of the current cell where the navigator is positioned.
        CellFormat cellformat = navigator.getCellFormat();

        // Clear all existing formatting applied to the current cell.
        cellformat.clearFormatting();

        // Set the background (shading) color of the cell.
        cellformat.getShading().setBackgroundPatternColor(Color.lightGray);

        // Save the modified document to the specified output file in DOCX format.
        doc.saveToFile("OperateTableCell.docx", FileFormat.Docx);

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

        // Set the background (shading) color of the table cell.
        cell.getCellFormat().getShading().setBackgroundPatternColor(Color.orange);

        // Vertically center the content within the table cell.
        cell.getCellFormat().setVerticalAlignment(VerticalAlignment.Middle);
    }
}
