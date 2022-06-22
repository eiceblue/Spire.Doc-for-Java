import com.spire.doc.*;
import com.spire.doc.documents.*;
import java.awt.*;

public class addTableByArray {
    public static void main(String[] args) {
         String output = "output/addTableByArray.docx";

        //create a Word document
        Document document = new Document();

        //add a section
        Section section = document.addSection();

        //add paragraph style
        ParagraphStyle style = new ParagraphStyle(document);
        style.getCharacterFormat().setFontSize(20f);
        style.getCharacterFormat().setBold(true);
        style.getCharacterFormat().setTextColor(Color.ORANGE);
        document.getStyles().add(style);

        //create a paragraph and append text
        Paragraph para = section.addParagraph();
        para.appendText("Table");

        //apply style for the added paragraph
        para.getFormat().setHorizontalAlignment(HorizontalAlignment.Center);
        para.applyStyle(style.getName());
        Table table =section.addTable(true);
        String[][] data =
                {
                        new String[]{"Name", "Capital", "Continent", "Area", "Population"},
                        new String[]{"Argentina", "Buenos Aires", "South America", "2777815", "32300003"},
                        new String[]{"Bolivia", "La Paz", "South America", "1098575", "7300000" },
                        new String[]{"Brazil", "Brasilia", "South America", "8511196", "150400000"},
                };


        int rowCount =data.length;
        int columnCount =data[0].length;
        table.resetCells(rowCount,columnCount);

        //fill the data to Table
        for (int i = 0; i <rowCount; i++)
        {
            for (int j = 0; j < columnCount; j++)
            {
                table.getRows().get(i).getCells().get(j).addParagraph().appendText(data[i][j]);
            }
        }
        //save the document
        document.saveToFile(output, FileFormat.Docx);
    }
}
