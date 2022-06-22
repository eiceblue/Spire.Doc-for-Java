import com.spire.doc.*;

public class removeTextBox {
    public static void main(String[] args) {
        //Load the document
        Document doc = new Document();
        doc.loadFromFile("data/textBoxTemplate.docx");

        //Remove the first text box
        doc.getTextBoxes().removeAt(0);

        //Clear all the text boxes
        //doc.getTextBoxes().clear();

        //Save the document
        String output = "output/removeTextBox.docx";
        doc.saveToFile(output, FileFormat.Docx);
    }
}
