import com.spire.doc.*;
import com.spire.doc.documents.*;

public class removeBookmarkContent {
    public static void main(String[] args) {
        String input = "data/bookmarks.docx";
        String output = "output/removeBookmarkContent.docx";

        //Load the document from disk.
        Document document = new Document();
        document.loadFromFile(input);

        //Get the bookmark by name.
        Bookmark bookmark = document.getBookmarks().get("Test2");

        //Get the BookmarkStart paragraph
        Paragraph para = (Paragraph)bookmark.getBookmarkStart().getOwner();

        //Get the start index
        int startIndex = para.getChildObjects().indexOf(bookmark.getBookmarkStart());

        //Get the BookmarkEnd paragraph
        para = (Paragraph)bookmark.getBookmarkEnd().getOwner();

        //Get the end index
        int endIndex = para.getChildObjects().indexOf(bookmark.getBookmarkEnd());

        //Remove the content of the bookmark.
        for (int i = startIndex + 1; i < endIndex; i++)
        {
            para.getChildObjects().removeAt(startIndex + 1);
        }
        //Save the document.
        document.saveToFile(output, FileFormat.Docx);
    }
}
