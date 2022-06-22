import com.spire.doc.*;
import com.spire.doc.documents.*;

public class replaceBookmarkContent {
    public static void main(String[] args) {
        String input = "data/bookmarks.docx";
        String output = "output/replaceBookmarkContent.docx";

        //Load the document from disk.
        Document doc = new Document();
        doc.loadFromFile(input);

        //Locate the bookmark.
        BookmarksNavigator bookmarkNavigator = new BookmarksNavigator(doc);
        bookmarkNavigator.moveToBookmark("Test2");

        //Replace the context with new.
        bookmarkNavigator.replaceBookmarkContent("This is replaced content.", false);

        //Save the document.
        doc.saveToFile(output, FileFormat.Docx);
    }
}
