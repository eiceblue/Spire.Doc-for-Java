import com.spire.doc.*;
import com.spire.doc.documents.*;
import com.spire.doc.fields.DocPicture;

public class setVerticalAlignment {
    public static void main(String[] args) {
        //Create a new Word document and add a new section
        Document doc = new Document();
        Section section = doc.addSection();

        //Add a table with 3 columns and 3 rows
        Table table = section.addTable(true);
        table.resetCells(3, 3);

        //Merge rows
        table.applyVerticalMerge(0, 0, 2);

        //Set the vertical alignment for each cell, default is top
        table.getRows().get(0).getCells().get(0).getCellFormat().setVerticalAlignment(VerticalAlignment.Middle);
        table.getRows().get(0).getCells().get(1).getCellFormat().setVerticalAlignment(VerticalAlignment.Top);
        table.getRows().get(0).getCells().get(2).getCellFormat().setVerticalAlignment(VerticalAlignment.Top);
        table.getRows().get(1).getCells().get(1).getCellFormat().setVerticalAlignment(VerticalAlignment.Middle);
        table.getRows().get(1).getCells().get(2).getCellFormat().setVerticalAlignment(VerticalAlignment.Middle);
        table.getRows().get(2).getCells().get(1).getCellFormat().setVerticalAlignment(VerticalAlignment.Bottom);
        table.getRows().get(2).getCells().get(2).getCellFormat().setVerticalAlignment(VerticalAlignment.Bottom);
        //Inset a picture into the table cell
        Paragraph paraPic = table.getRows().get(0).getCells().get(0).addParagraph();
        DocPicture pic = paraPic.appendPicture("data/E-iceblue.png");

        //Create data
        String[][] data = {
                new String[] {"","Spire.Office","Spire.DataExport"},
                new String[] {"","Spire.Doc","Spire.DocViewer"},
                new String[] {"","Spire.XLS","Spire.PDF"}
        };

        //Append data to table
        for (int r = 0; r < 3; r++)
        {
            TableRow dataRow = table.getRows().get(r);
            dataRow.setHeight(50f);
            for (int c = 0; c < 3; c++)
            {
                if (c == 1)
                {
                    Paragraph par = dataRow.getCells().get(c).addParagraph();
                    par.appendText(data[r][c]);
                    dataRow.getCells().get(c).setWidth((section.getPageSetup().getClientWidth()) / 2);
                }
                if (c == 2)
                {
                    Paragraph par = dataRow.getCells().get(c).addParagraph();
                    par.appendText(data[r][c]);
                    dataRow.getCells().get(c).setWidth((section.getPageSetup().getClientWidth()) / 2);
                }
            }
        }

        //Save and launch document
        String output = "output/SetVerticalAlignment.docx";
        doc.saveToFile(output, FileFormat.Docx_2013);
    }
}
