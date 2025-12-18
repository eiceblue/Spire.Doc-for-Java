import com.spire.doc.*;
import com.spire.doc.documents.*;

import java.awt.*;

public class createTableDirectly {
    public static void main(String[] args) {
        // Create a new document object
        Document doc = new Document();
        //Add a section
        Section section = doc.addSection();
        //Create a table
        Table table = new Table(doc);
        table.resetCells(1,2);

        //Set the width of table
        table.setPreferredWidth(new PreferredWidth(WidthType.Percentage, (short)100));
        //Set the border of table
        table.getFormat().getBorders().setBorderType( BorderStyle.Single);

        //Create a table row
        TableRow row = table.getRows().get(0);
        row.setHeight(50.0f);

        //Create a table cell
        TableCell cell1 = table.getRows().get(0).getCells().get(0);
        //Add a paragraph
        Paragraph para1 = cell1.addParagraph();
        //Append text in the paragraph
        para1.appendText("Row 1, Cell 1");
        //Set the horizontal alignment of paragrah
        para1.getFormat().setHorizontalAlignment(HorizontalAlignment.Center);
        //Set the background color of cell
        cell1.getCellFormat().getShading().setBackgroundPatternColor(Color.lightGray);
        //Set the vertical alignment of paragraph
        cell1.getCellFormat().setVerticalAlignment(VerticalAlignment.Middle);

        //Create a table cell
        TableCell cell2 = table.getRows().get(0).getCells().get(1);
        Paragraph para2 = cell2.addParagraph();
        para2.appendText("Row 1, Cell 2");
        para2.getFormat().setHorizontalAlignment(HorizontalAlignment.Center);
        cell2.getCellFormat().getShading().setBackgroundPatternColor(Color.lightGray);
        cell2.getCellFormat().setVerticalAlignment(VerticalAlignment.Middle);

        //Add the table in the section
        section.getTables().add(table);

        // Specify the output file path
        String output = "output/CreateTableDirectly_out.docx";

        // Save the document to the specified file format
        doc.saveToFile(output, FileFormat.Docx_2013);

        // Dispose the document resources
        doc.dispose();
    }
}
