import com.spire.doc.*;
import com.spire.doc.documents.*;
import com.spire.doc.fields.*;

public class addImageToEachPage {
    public static void main(String[] args) {
        String input1 = "data/multiPages.docx";
        String input2 = "data/spireDoc.png";
        String output = "output/addImageToEachPage.docx";

        //Open a Word document
        Document document = new Document();
        document.loadFromFile(input1);

        //Add a picture in footer and set it's position
        DocPicture picture = document.getSections().get(0).getHeadersFooters().getFooter().addParagraph().appendPicture(input2);
        picture.setVerticalOrigin( VerticalOrigin.Page);
        picture.setHorizontalOrigin( HorizontalOrigin.Page);
        picture.setVerticalAlignment(ShapeVerticalAlignment.Bottom);
        picture.setTextWrappingStyle(TextWrappingStyle.None);

        //Add a textbox in footer and set it's positiion
        TextBox textbox = document.getSections().get(0).getHeadersFooters().getFooter().addParagraph().appendTextBox(150, 20);
        textbox.setVerticalOrigin( VerticalOrigin.Page);
        textbox.setHorizontalOrigin( HorizontalOrigin.Page);
        textbox.setHorizontalPosition(300);
        textbox.setVerticalPosition(800);
        textbox.getBody().addParagraph().appendText("Welcome to E-iceblue");

        //Save the document
        document.saveToFile(output, FileFormat.Docx);
    }
}
