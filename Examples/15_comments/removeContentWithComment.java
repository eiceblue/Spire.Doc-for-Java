import com.spire.doc.*;
import com.spire.doc.documents.*;
import com.spire.doc.fields.*;
import java.util.ArrayList;

public class removeContentWithComment {
    public static void main(String[] args) {
        String input = "data/commentSample.docx";
        String output = "output/removeContentWithComment.docx";

        //Create a document
        Document document = new Document();

        //Load the document from disk.
        document.loadFromFile(input);

        //Get the first comment
        Comment comment = document.getComments().get(1);

        //Get the paragraph of obtained comment
        Paragraph para = comment.getOwnerParagraph();

        //Get index of the CommentMarkStart
        int startIndex = para.getChildObjects().indexOf(comment.getCommentMarkStart());

        //Get index of the CommentMarkEnd
        int endIndex = para.getChildObjects().indexOf(comment.getCommentMarkEnd());

        //Create a list
        ArrayList<TextRange> list = new ArrayList<TextRange>();

        //Get TextRanges between the indexes
        for (int i = startIndex; i < endIndex; i++)
        {
            if (para.getChildObjects().get(i) instanceof TextRange)
            {
                list.add((TextRange)para.getChildObjects().get(i) );
            }
        }

        //Insert a new TextRange
        TextRange textRange = new TextRange(document);

        //Set text is null
        textRange.setText(null);

        //Insert the new textRange
        para.getChildObjects().insert(endIndex, textRange);

        //Remove previous TextRanges
        for (int i = 0; i < list.size(); i++)
        {
            para.getChildObjects().remove(list.get(i));
        }
        //Save the document.
        document.saveToFile(output, FileFormat.Docx);
    }
}
