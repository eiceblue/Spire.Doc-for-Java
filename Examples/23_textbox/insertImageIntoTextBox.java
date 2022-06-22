import com.spire.doc.*;
import com.spire.doc.documents.*;
import com.spire.doc.fields.*;

public class insertImageIntoTextBox {
    public static void main(String[] args) {
        //Create a new document
        Document doc = new Document();

        Section section = doc.addSection();
        Paragraph paragraph = section.addParagraph();

        //Append a textbox to paragraph
        TextBox tb = paragraph.appendTextBox(220, 220);

        //Set the position of the textbox
        tb.getFormat().setHorizontalOrigin(HorizontalOrigin.Page);
        tb.getFormat().setHorizontalPosition(50);
        tb.getFormat().setVerticalOrigin(VerticalOrigin.Page);
        tb.getFormat().setVerticalPosition(50);

        //Set the fill effect of textbox as picture
        tb.getFormat().getFillEfects().setType(BackgroundType.Picture);

        //Fill the textbox with a picture
        tb.getFormat().getFillEfects().setPicture("data/spire.Doc.png");

        //Save the document
        String output = "output/insertImageIntoTextBox.docx";
        doc.saveToFile(output, FileFormat.Docx);
    }
}
