import com.spire.doc.*;
import com.spire.doc.fields.*;

public class replyToComment {
    public static void main(String[] args) {
        String input1 = "data/commentSample.docx";
        String input2 = "data/e-iceblue.png";
        String output = "output/replyToComment.docx";

        //Load the document from disk.
        Document doc = new Document();
        doc.loadFromFile(input1);

        //get the first comment.
        Comment comment1 = doc.getComments().get(0);

        //create a new comment and specify the author and content.
        Comment replyComment1 = new Comment(doc);
        replyComment1.getFormat().setAuthor("E-iceblue");
        replyComment1.getBody().addParagraph().appendText("Spire.Doc for Java is a professional Word Java library on operating Word documents.");

        //add the new comment as a reply to the selected comment.
        comment1.replyToComment(replyComment1);

        //insert a picture in the comment
        DocPicture docPicture = new DocPicture(doc);
        docPicture.loadImage(input2);
        replyComment1.getBody().getParagraphs().get(0).getChildObjects().add(docPicture);

        //Save the document.
        doc.saveToFile(output, FileFormat.Docx);
    }
}
