import com.spire.doc.*;
import com.spire.doc.fields.*;

public class addBarcodeImage {
    public static void main(String[] args) {
        String input1 = "data/sample.docx";
        String input2 = "data/barcode.png";
        String output = "output/addBarcodeImage.docx";

        //Open a Word document
        Document document = new Document();
        document.loadFromFile(input1);

        //Add barcode image
        DocPicture picture = document.getSections().get(0).addParagraph().appendPicture(input2);

        //Save the document
        document.saveToFile(output, FileFormat.Docx);
    }
}
