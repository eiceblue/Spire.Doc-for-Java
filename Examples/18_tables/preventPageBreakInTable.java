import com.spire.doc.*;
import com.spire.doc.documents.Paragraph;

public class preventPageBreakInTable {
    public static void main(String[] args) {
        //Create Word document.
        Document document = new Document();

        //Load the file from disk.
        document.loadFromFile("data/JAVATemplate_N.docx");

        //Get the table from Word document.
        Table table = document.getSections().get(0).getTables().get(0);

        //Change the paragraph setting to keep them together.
        for (TableRow row:(Iterable<TableRow>)table.getRows())
        {
            for (TableCell cell : (Iterable<TableCell>)row.getCells())
            {
                for (Paragraph p : (Iterable<Paragraph>)cell.getParagraphs())
                {
                    p.getFormat().setKeepFollow(true);
                }
            }
        }

        //Save to file.
        String result = "output/Result-PreventPageBreaksInWordTable.docx";
        document.saveToFile(result, FileFormat.Docx_2013);
    }
}
