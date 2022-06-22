import com.spire.doc.*;
import com.spire.doc.documents.*;

public class insertDocAtBookmark {
    public static void main(String[] args) {
        String input1 = "data/bookmarks.docx";
        String input2 = "data/insert-LL.docx";
        String output = "output/insertDocAtBookmark.docx";

        //Create the first document
        Document document1 = new Document();

        //Load the first document from disk.
        document1.loadFromFile(input1);

        //Create the second document
        Document document2 = new Document();

        //Load the second document from disk.
        document2.loadFromFile(input2);

        //Get the first section of the first document
        Section section1 = document1.getSections().get(0);

        //Locate the bookmark
        BookmarksNavigator bn = new BookmarksNavigator(document1);

        //Find bookmark by name
        bn.moveToBookmark("Test2", true, true);

        //Get bookmarkStart
        BookmarkStart start = bn.getCurrentBookmark().getBookmarkStart();

        //Get the owner paragraph
        Paragraph para = start.getOwnerParagraph();

        //Get the para index
        int index = section1.getBody().getChildObjects().indexOf(para);

        //Insert the paragraphs of document2
        for (int i = 0; i < document2.getSections().getCount(); i++)
        {
            for (int j = 0; j < document2.getSections().get(i).getParagraphs().getCount(); j++)
            {
                Paragraph insertPara= (Paragraph)document2.getSections().get(i).getParagraphs().get(j).deepClone();
                section1.getBody().getChildObjects().insert(index++ + 1, insertPara);
            }
        }
        //Save the document.
        document1.saveToFile(output, FileFormat.Docx);
    }
}
