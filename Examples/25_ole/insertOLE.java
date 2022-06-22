import com.spire.doc.*;
import com.spire.doc.documents.*;
import com.spire.doc.fields.*;

public class insertOLE {
    public static void main(String[] args) {
        //Create a document
        Document doc = new Document();

        //Add a section
        Section sec = doc.addSection();

        //Add a paragraph
        Paragraph par = sec.addParagraph();

        //Load the image
        DocPicture picture = new DocPicture(doc);
        picture.loadImage("data/excel.png");

        //Insert the OLE
        DocOleObject obj = par.appendOleObject("data/example.xlsx", picture, OleObjectType.Excel_Worksheet);

        //Save to file
        String output = "output/insertOLE.docx";
        doc.saveToFile(output, FileFormat.Docx_2013);
    }
}
