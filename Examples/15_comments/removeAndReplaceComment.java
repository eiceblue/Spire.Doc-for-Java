import com.spire.doc.*;

public class removeAndReplaceComment {
    public static void main(String[] args) {
        String input = "data/commentSample.docx";
        String output = "output/removeAndReplaceComment.docx";

        //Load the document
        Document doc = new Document();
        doc.loadFromFile(input);

        //Replace the content of the first comment
        doc.getComments().get(0).getBody().getParagraphs().get(0).replace("This is the title", "This comment is changed.", false, false);

        //Remove the second comment
        doc.getComments().removeAt(1);

        //Save the document
        doc.saveToFile(output, FileFormat.Docx);
    }
}
