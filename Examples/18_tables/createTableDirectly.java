import com.spire.doc.*;
import com.spire.doc.documents.*;

import java.awt.*;

public class createTableDirectly {
    public static void main(String[] args) {
        //Create a Word document
        Document doc = new Document();

        //Add a section
        Section section = doc.addSection();

        //Create a table
        Table table = new Table(doc);
        //Set the width of table
        table.setPreferredWidth(new PreferredWidth(WidthType.Percentage, (short)100));
        //Set the border of table
        table.getTableFormat().getBorders().setBorderType( BorderStyle.Single);

        //Create a table row
        TableRow row = new TableRow(doc);
        row.setHeight(50.0f);
        table.getRows().add(row);

        //Create a table cell
        TableCell cell1 = new TableCell(doc);
        //Add a paragraph
        Paragraph para1 = cell1.addParagraph();
        //Append text in the paragraph
        para1.appendText("Row 1, Cell 1");
        //Set the horizontal alignment of paragrah
        para1.getFormat().setHorizontalAlignment(HorizontalAlignment.Center);
        //Set the background color of cell
        cell1.getCellFormat().setBackColor(Color.lightGray);
        //Set the vertical alignment of paragraph
        cell1.getCellFormat().setVerticalAlignment(VerticalAlignment.Middle);
        row.getCells().add(cell1);

        //Create a table cell
        TableCell cell2 = new TableCell(doc);
        Paragraph para2 = cell2.addParagraph();
        para2.appendText("Row 1, Cell 2");
        para2.getFormat().setHorizontalAlignment(HorizontalAlignment.Center);
        cell2.getCellFormat().setBackColor(Color.lightGray);
        cell2.getCellFormat().setVerticalAlignment(VerticalAlignment.Middle);
        row.getCells().add(cell2);

        //Add the table in the section
        section.getTables().add(table);

        //Save the document
        String output = "output/CreateTableDirectly_out.docx";
        doc.saveToFile(output, FileFormat.Docx_2013);
    }
}
