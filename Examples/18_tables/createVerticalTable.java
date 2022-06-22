import com.spire.doc.*;
import com.spire.doc.documents.TextDirection;

public class createVerticalTable {
    public static void main(String[] args) {
        //Create Word document.
        Document document = new Document();

        //Add a new section.
        Section section = document.addSection();

        //Add a table with rows and columns and set the text for the table.
        Table table = section.addTable();
        table.resetCells(1, 1);
        TableCell cell = table.getRows().get(0).getCells().get(0);
        table.getRows().get(0).setHeight(150f);
        cell.addParagraph().appendText("Draft copy in vertical style");

        //Set the TextDirection for the table to RightToLeftRotated.
        cell.getCellFormat().setTextDirection(TextDirection.Right_To_Left_Rotated);

        //Set the table format.
        table.getTableFormat().setWrapTextAround(true);
        table.getTableFormat().getPositioning().setVertRelationTo(VerticalRelation.Page);
        table.getTableFormat().getPositioning().setHorizRelationTo( HorizontalRelation.Page);
        table.getTableFormat().getPositioning().setHorizPosition((float) section.getPageSetup().getPageSize().getWidth()-table.getWidth());
        table.getTableFormat().getPositioning().setVertPosition( 200f);

        //Save to file.
        String result = "output/Result-CreateVerticalTable.docx";
        document.saveToFile(result, FileFormat.Docx_2013);
    }
}
