import com.spire.doc.*;
import com.spire.doc.documents.*;
import com.spire.doc.fields.*;

import java.awt.*;

public class textBoxFormat {
    public static void main(String[] args) {
        //Create a new document
        Document doc = new Document();
        Section sec = doc.addSection();

        //Add a text box and append sample text
        TextBox TB = doc.getSections().get(0).addParagraph().appendTextBox(310, 90);
        Paragraph para = TB.getBody().addParagraph();
        TextRange TR = para.appendText("Using Spire.Doc, developers will find " +
                "a simple and effective method to endow their applications with rich MS Word features. ");
        TR.getCharacterFormat().setFontName("Cambria ");
        TR.getCharacterFormat().setFontSize(13);

        //Set exact position for the text box
        TB.getFormat().setHorizontalOrigin(HorizontalOrigin.Page);
        TB.getFormat().setHorizontalPosition(120);
        TB.getFormat().setVerticalOrigin(VerticalOrigin.Page);
        TB.getFormat().setVerticalPosition(100);

        //Set line style for the text box
        TB.getFormat().setLineStyle(TextBoxLineStyle.Double);
        TB.getFormat().setLineColor(Color.BLUE);
        TB.getFormat().setLineDashing(LineDashing.Solid);
        TB.getFormat().setLineWidth(5);

        //Set internal margin for the text box
        TB.getFormat().getInternalMargin().setTop(15);
        TB.getFormat().getInternalMargin().setBottom(10);
        TB.getFormat().getInternalMargin().setLeft(12);
        TB.getFormat().getInternalMargin().setRight(10);

        //Save the document
        String output = "output/textBoxFormat.docx";
        doc.saveToFile(output, FileFormat.Docx_2013);
    }
}
