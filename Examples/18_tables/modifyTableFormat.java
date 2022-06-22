import com.spire.doc.*;
import com.spire.doc.documents.*;

import java.awt.*;

public class modifyTableFormat {
    public static void main(String[] args) {
        //Load Word document from disk
        Document document = new Document();
        document.loadFromFile("data/ModifyTableFormat.docx");

        //Get the first section
        Section section = document.getSections().get(0);

        //Get tables
        Table tb1 = section.getTables().get(0);
        Table tb2 = section.getTables().get(1);
        Table tb3 = section.getTables().get(2);

        MoidyTableFormat(tb1);
        ModifyRowFormat(tb2);
        ModifyCellFormat(tb3);

        String output = "output/ModifyTableFormat_out.docx";
        document.saveToFile(output, FileFormat.Docx_2013);
    }
    private static void MoidyTableFormat(Table table)
    {
        //Set table width
        table.setPreferredWidth(new PreferredWidth(WidthType.Twip,(short)6000));

        //Apply style for table
        table.applyStyle(DefaultTableStyle.Table_3_Deffects_2);

        //Set table padding
        table.getTableFormat().getPaddings().setAll(5f);

        //Set table title and description
        table.setTitle( "Spire.Doc for Java");
        table.setTableDescription("Spire.Doc for Java is a professional Word Java library");
    }
    private static void ModifyRowFormat(Table table)
    {
        //Set cell spacing
        table.getRows().get(0).getRowFormat().setCellSpacing(2f);

        //Set row height
        table.getRows().get(1).setHeightType( TableRowHeightType.Exactly);
        table.getRows().get(1).setHeight(20f);

        //Set background color
        table.getRows().get(2).getRowFormat().setBackColor(Color.gray);
    }
    private static void ModifyCellFormat(Table table)
    {
        //Set alignment
        table.getRows().get(0).getCells().get(0).getCellFormat().setVerticalAlignment(VerticalAlignment.Middle);
        table.getRows().get(0).getCells().get(0).getParagraphs().get(0).getFormat().setHorizontalAlignment(HorizontalAlignment.Center);

        //Set background color
        table.getRows().get(1).getCells().get(0).getCellFormat().setBackColor(Color.gray);

        //Set cell border
        table.getRows().get(2).getCells().get(0).getCellFormat().getBorders().setBorderType(BorderStyle.Single);
        table.getRows().get(2).getCells().get(0).getCellFormat().getBorders().setLineWidth(1f);
        table.getRows().get(2).getCells().get(0).getCellFormat().getBorders().getLeft().setColor(Color.red);
        table.getRows().get(2).getCells().get(0).getCellFormat().getBorders().getRight().setColor(Color.red);
        table.getRows().get(2).getCells().get(0).getCellFormat().getBorders().getTop().setColor(Color.red);
        table.getRows().get(2).getCells().get(0).getCellFormat().getBorders().getBottom().setColor(Color.red);

        //Set text direction
        table.getRows().get(3).getCells().get(0).getCellFormat().setTextDirection(TextDirection.Right_To_Left);
    }
}
