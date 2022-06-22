package comments;
import com.spire.doc.*;
import com.spire.doc.documents.*;
import com.spire.doc.fields.*;

public class comment {
    public static void main(String[] args) {
        String input = "data/sample.docx";
        String output = "output/comment.docx";

        //Load the document from disk.
        Document document = new Document();
        document.loadFromFile(input);

        //Insert comments
        InsertComments(document.getSections().get(0));

        //Save the document.
        document.saveToFile(output, FileFormat.Docx);
    }
    private static void InsertComments(Section section)
    {
        //Insert comment.
        Paragraph paragraph = section.getParagraphs().get(1);
        Comment comment = paragraph.appendComment("Spire.Doc for java");
        comment.getFormat().setAuthor("E-iceblue");
        comment.getFormat().setInitial("CM");
    }
}
