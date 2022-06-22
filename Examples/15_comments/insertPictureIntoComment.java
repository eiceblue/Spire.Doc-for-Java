import com.spire.doc.*;
import com.spire.doc.documents.*;
import com.spire.doc.fields.*;

public class insertPictureIntoComment {
    public static void main(String[] args) {
        String input1 = "data/sample.docx";
        String input2 = "data/e-iceblue.png";
        String output = "output/insertPictureIntoComment.docx";

        //Load the document
        Document doc = new Document();
        doc.loadFromFile(input1);

        //Get the first paragraph and insert comment
        Paragraph paragraph = doc.getSections().get(0).getParagraphs().get(2);
        Comment comment = paragraph.appendComment("This is a comment.");
        comment.getFormat().setAuthor("E-iceblue");

        //Load a picture
        DocPicture docPicture = new DocPicture(doc);
        docPicture.loadImage(input2);

        //Insert the picture into the comment body
        comment.getBody().addParagraph().getChildObjects().add(docPicture);

        //Save the document
        doc.saveToFile(output, FileFormat.Docx);
    }
}
