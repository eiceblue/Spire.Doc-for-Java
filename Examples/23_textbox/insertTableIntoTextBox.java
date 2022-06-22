import com.spire.doc.*;
import com.spire.doc.documents.*;
import com.spire.doc.fields.*;

public class insertTableIntoTextBox {
    public static void main(String[] args) {
        //Create a new document
        Document doc = new Document();

        //Add a section
        Section section = doc.addSection();

        //Add a paragraph to the section
        Paragraph paragraph = section.addParagraph();

        //Add a textbox to the paragraph
        TextBox textbox = paragraph.appendTextBox(300, 100);

        //Set the position of the textbox
        textbox.getFormat().setHorizontalOrigin(HorizontalOrigin.Page);
        textbox.getFormat().setHorizontalPosition(140);
        textbox.getFormat().setVerticalOrigin(VerticalOrigin.Page);
        textbox.getFormat().setVerticalPosition(50);

        //Add text to the textbox
        Paragraph textboxParagraph = textbox.getBody().addParagraph();
        TextRange textboxRange = textboxParagraph.appendText("Table 1");
        textboxRange.getCharacterFormat().setFontName("Arial");

        //Insert table to the textbox
        Table table = textbox.getBody().addTable(true);

        //Specify the number of rows and columns of the table
        table.resetCells(4, 4);

        String[][] data = new String[][]
                {
                        {"Name", "Age", "Gender", "ID"},
                        {"John", "28", "Male", "0023"},
                        {"Steve", "30", "Male", "0024"},
                        {"Lucy", "26", "female", "0025"}
                };

        //Add data to the table
        for (int i = 0; i < 4; i++) {
            for (int j = 0; j < 4; j++) {

                TextRange tableRange = table.getRows().get(i).getCells().get(j).addParagraph().appendText(data[i][j]);
                tableRange.getCharacterFormat().setFontName("Arial");
            }
        }

        //Apply style to the table
        table.applyStyle(DefaultTableStyle.Table_Colorful_2);

        //Save the document
        String output = "output/insertTableIntoTextBox.docx";
        doc.saveToFile(output, FileFormat.Docx_2013);
    }
}
