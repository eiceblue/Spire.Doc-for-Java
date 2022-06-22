import com.spire.doc.*;
import com.spire.doc.documents.*;
import com.spire.doc.fields.*;

public class addCommentForSpecificText {
    public static void main(String[] args) {
        //Create word document
        Document document = new Document();

        //Load the document from disk
        document.loadFromFile("data/commentTemplate.docx");

        InsertComments(document, "Development");

        //Save the document
        document.saveToFile("output/addCommentForTextRange.docx", FileFormat.Docx);
    }

    private static void InsertComments(Document doc, String keystring) {
        //Find the key string
        TextSelection find = doc.findString(keystring, false, true);

        //Create the commentMarkStart and commentMarkEnd
        CommentMark commentMarkStart = new CommentMark(doc);
        commentMarkStart.setType(CommentMarkType.Comment_Start);
        CommentMark commentMarkEnd = new CommentMark(doc);
        commentMarkEnd.setType(CommentMarkType.Comment_End);

        //Add the content for comment
        Comment comment = new Comment(doc);
        comment.getBody().addParagraph().setText("Test comments");
        comment.getFormat().setAuthor("E-iceblue");

        //Get the textRange
        TextRange range = find.getAsOneRange();

        //Get its paragraph
        Paragraph para = range.getOwnerParagraph();

        //Get the index of textRange
        int index = para.getChildObjects().indexOf(range);

        //Add comment
        para.getChildObjects().add(comment);

        //Insert the commentMarkStart and commentMarkEnd
        para.getChildObjects().insert(index, commentMarkStart);
        para.getChildObjects().insert(index + 2, commentMarkEnd);
    }
}
