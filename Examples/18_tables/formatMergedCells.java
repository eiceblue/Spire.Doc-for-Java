import com.spire.doc.*;
import com.spire.doc.documents.*;
import com.spire.doc.fields.*;

import java.awt.*;

public class formatMergedCells {
    public static void main(String[] args) {
        //Create word document
        Document document = new Document();

        //Add a new section
        Section section = document.addSection();

        //Add a new table
        Table table = addTable(section);

        //Create a new style
        ParagraphStyle style = new ParagraphStyle(document);
        style.setName("Style");
        style.getCharacterFormat().setTextColor(Color.cyan);
        style.getCharacterFormat().setItalic(true);
        style.getCharacterFormat().setBold(true);
        style.getCharacterFormat().setFontSize(13);
        document.getStyles().add(style);

        //Merge cell horizontally
        table.applyHorizontalMerge(0, 0, 1);
        //Apply style
        table.get(0, 0).getParagraphs().get(0).applyStyle(style.getName());
        //Set vertical and horizontal alignment
        table.get(0, 0).getCellFormat().setVerticalAlignment(VerticalAlignment.Middle);
        table.get(0, 0).getParagraphs().get(0).getFormat().setHorizontalAlignment(HorizontalAlignment.Center);

        //Merge cell vertically
        table.applyVerticalMerge(0, 1, 3);
        //Apply style
        table.get(1, 0).getParagraphs().get(0).applyStyle(style.getName());
        //Set vertical and horizontal alignment
        table.get(1, 0).getCellFormat().setVerticalAlignment(VerticalAlignment.Middle);
        table.get(1, 0).getParagraphs().get(0).getFormat().setHorizontalAlignment(HorizontalAlignment.Left);

        //Set column width
        table.get(1, 0).setCellWidth(20, CellWidthType.Percentage);

        //Save the file
        String output = "output/formatMergedCells.docx";
        document.saveToFile(output, FileFormat.Docx);
    }

    private static Table addTable(Section section) {
        Table table = section.addTable(true);
        table.resetCells(4, 3);
        //Table data
        String[][] data =
                {
                        new String[]{"Product", "", "Price"},
                        new String[]{"Spire.Doc", "Pro Edition", "$799"},
                        new String[]{"", "Standard Edition", "$599"},
                        new String[]{"", "Free Edition", "$0"},
                };

        for (int r = 0; r < data.length; r++) {
            TableRow dataRow = table.getRows().get(r);
            dataRow.setHeight(20);
            dataRow.setHeightType(TableRowHeightType.Exactly);
            dataRow.getRowFormat().setBackColor(Color.white);
            for (int c = 0; c < dataRow.getCells().getCount(); c++) {
                if (data[r][c] != null && data[r][c] != "") {
                    TextRange range = dataRow.getCells().get(c).addParagraph().appendText(data[r][c]);
                    range.getCharacterFormat().setFontName("Arial");
                }
            }
        }

        return table;
    }
}
