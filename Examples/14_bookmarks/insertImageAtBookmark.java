import com.spire.doc.*;
import com.spire.doc.documents.*;
import com.spire.doc.fields.*;

public class insertImageAtBookmark {
    public static void main(String[] args) {
        String input1 = "data/bookmarks.docx";
        String input2 = "data/e-iceblue.png";
        String output = "output/insertImageAtBookmark.docx";

        //Load the document
        Document doc = new Document();
        doc.loadFromFile(input1);

        //Create an instance of BookmarksNavigator
        BookmarksNavigator bn = new BookmarksNavigator(doc);

        //Find a bookmark named Test
        bn.moveToBookmark("Test2", true, true);

        //Add a section
        Section section0 = doc.addSection();

        //Add a paragraph for the section
        Paragraph paragraph = section0.addParagraph();

        //Add a picture into the paragraph
        DocPicture picture = paragraph.appendPicture(input2);

        //Add the paragraph at the position of bookmark
        bn.insertParagraph(paragraph);

        //Remove the section0
        doc.getSections().remove(section0);

        //Save the document.
        doc.saveToFile(output, FileFormat.Docx);
    }
}
