import com.spire.doc.*;
import com.spire.doc.fields.*;

public class deleteTableFromTextBox {
    public static void main(String[] args) {
        //Load the document
        Document doc = new Document();
        doc.loadFromFile("data/textBoxTable.docx");

        //Get the first textbox
        TextBox textbox = doc.getTextBoxes().get(0);

        //Remove the first table from the textbox
        textbox.getBody().getTables().removeAt(0);

        //Save the document
        String output = "output/deleteTableFromTextBox.docx";
        doc.saveToFile(output, FileFormat.Docx_2013);
    }
}
