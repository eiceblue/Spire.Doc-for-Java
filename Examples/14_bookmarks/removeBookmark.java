import com.spire.doc.*;

public class removeBookmark {
    public static void main(String[] args) {
        String input = "data/bookmarks.docx";
        String output = "output/removeBookmark.docx";

        //Load the document from disk.
        Document document = new Document();
        document.loadFromFile(input);

        //Get the bookmark by name.
        Bookmark bookmark = document.getBookmarks().get("Test2");

        //Remove the bookmark, not its content.
        document.getBookmarks().remove(bookmark);

        //Save the document.
        document.saveToFile(output, FileFormat.Docx);
    }
}
