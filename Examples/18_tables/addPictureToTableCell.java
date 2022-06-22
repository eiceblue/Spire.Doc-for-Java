import com.spire.doc.*;
import com.spire.doc.fields.*;

public class addPictureToTableCell {
    public static void main(String[] args) {
        String input1 = "data/tableTemplate.docx";
        String input2 = "data/spireDoc.png";
        String output = "output/addPictureToTableCell.docx";

        //load the document
        Document doc = new Document();
        doc.loadFromFile(input1);

        //get the first table from the first section of the document
        Table table1 = doc.getSections().get(0).getTables().get(0);

        //add a picture to the specified table cell and set picture size
        DocPicture picture = table1.getRows().get(1).getCells().get(2).getParagraphs().get(0).appendPicture(input2);
        picture.setWidth(100);
        picture.setHeight(100);

        //save the document
        doc.saveToFile(output, FileFormat.Docx);
    }
}
