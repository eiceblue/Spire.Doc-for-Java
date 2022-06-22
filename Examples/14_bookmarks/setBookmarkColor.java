import com.spire.doc.*;
import java.awt.*;

public class setBookmarkColor {
    public static void main(String[] args) {
        String input = "data/bookmarkTemplate.docx";
        String output = "output/setBookmarkColor.pdf";

        //Load the document
        Document doc = new Document();
        doc.loadFromFile(input);

        //Create an instance of ToPdfParameterList
        ToPdfParameterList toPdf = new ToPdfParameterList();

        //Set CreateWordBookmarks to true to use word bookmarks when create the bookmarks
        toPdf.setCreateWordBookmarks(true);

        //Set the title of word bookmarks
        toPdf.setWordBookmarksTitle("Changed bookmark");

        //Set the text color of word bookmarks
        toPdf.setWordBookmarksColor(Color.gray);

        //Call the event document_BookmarkLayout when drawing a bookmark
        doc.BookmarkLayout = new  com.spire.doc.documents.rendering.BookmarkLevelHandler() {
            @Override
            public void invoke(Object sender, com.spire.doc.documents.rendering.BookmarkLevelEventArgs args) {
                document_BookmarkLayout(sender, args);
            }
        };

        //Save the document
        doc.saveToFile(output, toPdf);
    }
    private static void document_BookmarkLayout(Object sender, com.spire.doc.documents.rendering.BookmarkLevelEventArgs args)
    {
        //set the different color for different levels of bookmarks
        if (args.getBookmarkLevel().getLevel() == 2)
        {
            args.getBookmarkLevel().setColor( Color.red);
            args.getBookmarkLevel().setStyle(BookmarkTextStyle.Bold);
        }
        else if (args.getBookmarkLevel().getLevel() == 3)
        {
            args.getBookmarkLevel().setColor( Color.gray);
            args.getBookmarkLevel().setStyle(BookmarkTextStyle.Italic);
        }
        else
        {
            args.getBookmarkLevel().setColor( Color.green);
            args.getBookmarkLevel().setStyle(BookmarkTextStyle.Regular);
        }
    }
}
