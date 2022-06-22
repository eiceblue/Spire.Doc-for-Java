import com.spire.doc.*;
import com.spire.doc.documents.*;

import java.awt.*;

public class repeatRowOnEachPage {
    public static void main(String[] args) {
        //Create word document
        Document document = new Document();

        //Create a new section
        Section section = document.addSection();

        //Create a table width default borders
        Table table = section.addTable(true);
        //Set table with to 100%
        PreferredWidth width = new PreferredWidth(WidthType.Percentage, (short)100);
        table.setPreferredWidth(width);

        //Add a new row
        TableRow row = table.addRow();
        //Set the row as a table header
        row.isHeader(true);
        //Set the backcolor of row
        row.getRowFormat().setBackColor(Color.lightGray);
        //Add a new cell for row
        TableCell cell = row.addCell();
        cell.setCellWidth(100,CellWidthType.Percentage);
        //Add a paragraph for cell to put some data
        Paragraph parapraph = cell.addParagraph();
        //Add text
        parapraph.appendText("Row Header 1");
        //Set paragraph horizontal center alignment
        parapraph.getFormat().setHorizontalAlignment(HorizontalAlignment.Center);

        row = table.addRow(false, 1);
        row.isHeader(true);
        row.getRowFormat().setBackColor(Color.orange);
        //Set row height
        row.setHeight(30f);
        cell = row.getCells().get(0);
        cell.setCellWidth(100,CellWidthType.Percentage);
        //Set cell vertical middle alignment
        cell.getCellFormat().setVerticalAlignment(VerticalAlignment.Middle);
        //Add a paragraph for cell to put some data
        parapraph = cell.addParagraph();
        //Add text
        parapraph.appendText("Row Header 2");
        parapraph.getFormat().setHorizontalAlignment(HorizontalAlignment.Center);

        //Add many common rows
        for (int i = 0; i < 70; i++)
        {
            row = table.addRow(false,2);
            cell = row.getCells().get(0);
            //Set cell width to 50% of table width
            cell.setCellWidth(50f,CellWidthType.Percentage);
            cell.addParagraph().appendText("Column 1 Text");
            cell = row.getCells().get(1);
            cell.setCellWidth(50f,CellWidthType.Percentage);
            cell.addParagraph().appendText("Column 2 Text");
        }
        //Set cell backcolor
        for (int j = 1; j < table.getRows().getCount(); j++)
        {
            if (j % 2 == 0)
            {
                TableRow row2 = table.getRows().get(j);
                for (int f = 0; f < row2.getCells().getCount(); f++)
                {
                    row2.getCells().get(f).getCellFormat().setBackColor(Color.PINK);
                }
            }
        }
        //Save file.
        String result = "output/RepeatRowOnEachPage_out.docx";
        document.saveToFile(result,FileFormat.Docx);
    }
}
