import com.spire.doc.*;
import com.spire.doc.documents.*;
import java.awt.*;

public class setDiagonalBorder {
    public static void main(String[] args) {
        String output = "output/setDiagonalBorder.docx";

        //create a document
        Document document = new Document();

        //add section
        Section section = document.addSection();

        //add a teable
        Table table = section.addTable(true);

        //set rows and columns
        table.resetCells(4, 3);

        //set table format
        table.getTableFormat().setHorizontalAlignment(RowAlignment.Center);

        //set double diagonal border
        table.getFirstRow().getCells().get(0).getCellFormat().getBorders().getDiagonalDown().setBorderType(BorderStyle.Double);
        table.getFirstRow().getCells().get(0).getCellFormat().getBorders().getDiagonalDown().setColor(Color.GREEN);
        table.getFirstRow().getCells().get(0).getCellFormat().getBorders().getDiagonalDown().setLineWidth(2f);

        //set single diagonal border
        table.getLastCell().getCellFormat().getBorders().getDiagonalUp().setBorderType(BorderStyle.Single);
        table.getLastCell().getCellFormat().getBorders().getDiagonalUp().setColor(Color.RED);
        table.getLastCell().getCellFormat().getBorders().getDiagonalUp().setLineWidth(0.8f);

        //save to file
        document.saveToFile(output, FileFormat.Docx);
    }
}
