import com.spire.doc.*;
import com.spire.doc.documents.*;
import com.spire.doc.fields.*;

public class createBookmarkForTable {
    public static void main(String[] args) {
        String output = "output/createBookmarkForTable.docx";

        //Create word document.
        Document document = new Document();

        //Add a new section.
        Section section = document.addSection();

        //Create bookmark for a table
        createBookmarkForTable(document, section);

        //Save the document.
        document.saveToFile(output, FileFormat.Docx);
    }
    private static void createBookmarkForTable(Document doc, Section section)
    {
        //Add a paragraph
        Paragraph paragraph = section.addParagraph();

        //Append text for added paragraph
        TextRange txtRange = paragraph.appendText("The following example demonstrates how to create bookmark for a table in a Word document.");

        //Set the font in italic
        txtRange.getCharacterFormat().setItalic(true);

        //Append bookmark start
        paragraph.appendBookmarkStart("CreateBookmark");

        //Append bookmark end
        paragraph.appendBookmarkEnd("CreateBookmark");

        //Add table
        Table table = section.addTable(true);

        //Set the number of rows and columns
        table.resetCells(2, 2);

        //Append text for table cells
        TextRange range = table.getRows().get(0).getCells().get(0).addParagraph().appendText("sampleA");
        range = table.getRows().get(0).getCells().get(1).addParagraph().appendText("sampleB");
        range = table.getRows().get(1).getCells().get(0).addParagraph().appendText("120");
        range = table.getRows().get(1).getCells().get(1).addParagraph().appendText("260");

        //Get the bookmark by index.
        Bookmark bookmark = doc.getBookmarks().get(0);

        //Get the name of bookmark.
        String bookmarkName = bookmark.getName();

        //Locate the bookmark by name.
        BookmarksNavigator navigator = new BookmarksNavigator(doc);
        navigator.moveToBookmark(bookmarkName);

        //Add table to TextBodyPart
        TextBodyPart part = navigator.getBookmarkContent();
        part.getBodyItems().add(table);

        //Replace bookmark cotent with table
        navigator.replaceBookmarkContent(part);
    }
}
