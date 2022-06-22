import com.spire.doc.*;
import com.spire.doc.documents.*;
import com.spire.doc.fields.DocPicture;

public class setOutsidePosition {
    public static void main(String[] args) {
        //Create a new word document and add a new section
        Document doc = new Document();
        Section sec = doc.addSection();

        //Get header
        HeaderFooter header = sec.getHeadersFooters().getHeader();

        //Add new paragraph on header and set HorizontalAlignment of the paragraph as left
        Paragraph paragraph = header.addParagraph();
        paragraph.getFormat().setHorizontalAlignment(HorizontalAlignment.Left);

        //Load an image for the paragraph
        //DocPicture headerimage = paragraph.appendPicture(Image.FromFile("data/Word.png"));
        DocPicture headerimage = paragraph.appendPicture("data/Word.png");
        //Add a table of 4 rows and 2 columns
        Table table = header.addTable();
        table.resetCells(4, 2);

        //Set the position of the table to the right of the image
        table.getTableFormat().setWrapTextAround(true);
        table.getTableFormat().getPositioning().setHorizPositionAbs(HorizontalPosition.Outside);
        table.getTableFormat().getPositioning().setVertRelationTo(VerticalRelation.Margin);
        table.getTableFormat().getPositioning().setVertPosition(43f);

        //Add contents for the table
        String[][] data = {
                new String[] {"Spire.Doc.left","Spire XLS.right"},
                new String[] {"Spire.Presentatio.left","Spire.PDF.right"},
                new String[] {"Spire.DataExport.left","Spire.PDFViewe.right"},
                new String []{"Spire.DocViewer.left","Spire.BarCode.right"}
        };

        for (int r = 0; r < 4; r++)
        {
            TableRow dataRow = table.getRows().get(r);
            for (int c = 0; c < 2; c++)
            {
                if (c == 0)
                {
                    Paragraph par = dataRow.getCells().get(c).addParagraph();
                    par.appendText(data[r][c]);
                    par.getFormat().setHorizontalAlignment(HorizontalAlignment.Left);
                    dataRow.getCells().get(c).setWidth(180f);
                }
                else
                {
                    Paragraph par = dataRow.getCells().get(c).addParagraph();
                    par.appendText(data[r][c]);
                    par.getFormat().setHorizontalAlignment(HorizontalAlignment.Right);
                    dataRow.getCells().get(c).setWidth(180f);
                }
            }
        }

        //Save and launch document
        String output = "output/SetOutsidePosition.docx";
        doc.saveToFile(output, FileFormat.Docx);
    }
}
