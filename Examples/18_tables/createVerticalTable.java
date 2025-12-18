import com.spire.doc.*;
import com.spire.doc.documents.TextDirection;

public class createVerticalTable {
    public static void main(String[] args) {
        // Create a new document object
        Document document = new Document();

        // Add a section to the document
        Section section = document.addSection();

        // Add a table to the section
        Table table = section.addTable();

        // Reset the table with 1 row and 1 column
        table.resetCells(1, 1);

        // Get the cell in the first row and first column
        TableCell cell = table.getRows().get(0).getCells().get(0);

        // Set the height of the first row to 150f (float)
        table.getRows().get(0).setHeight(150f);

        // Add a paragraph with text to the cell
        cell.addParagraph().appendText("Draft copy in vertical style");

        // Set the text direction to right-to-left rotated
        cell.getCellFormat().setTextDirection(TextDirection.Right_To_Left_Rotated);

        //Set the table format.
        table.getFormat().setWrapTextAround(true);
        table.getFormat().getPositioning().setVertRelationTo(VerticalRelation.Page);
        table.getFormat().getPositioning().setHorizRelationTo( HorizontalRelation.Page);
        table.getFormat().getPositioning().setHorizPosition((float) section.getPageSetup().getPageSize().getWidth()-table.getWidth());
        table.getFormat().getPositioning().setVertPosition( 200f);

        // Specify the output file path
        String result = "output/Result-CreateVerticalTable.docx";

        // Save the document to the specified file format
        document.saveToFile(result, FileFormat.Docx);

        // Dispose the document resources
        document.dispose();
    }
}
