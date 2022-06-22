import com.spire.doc.*;
import com.spire.doc.documents.*;
import com.spire.doc.fields.*;

public class replaceTextWithTable {
    public static void main(String[] args) {
        //Create word document.
        Document document = new Document();

        // Load the file from disk.
        document.loadFromFile("data/Template_Docx_1.docx");

        //Return TextSection by finding the key text string "Christmas Day, December 25".
        Section section = document.getSections().get(0);
        TextSelection selection = document.findString("Christmas Day, December 25", true, true);

        //Return TextRange from TextSection, then get OwnerParagraph through TextRange.
        TextRange range = selection.getAsOneRange();
        Paragraph paragraph = range.getOwnerParagraph();

        //Return the zero-based index of the specified paragraph.
        Body body = paragraph.ownerTextBody();
        int index = body.getChildObjects().indexOf(paragraph);

        //Create a new table.
        Table table = section.addTable(true);
        table.resetCells(3, 3);

        // Remove the paragraph and insert table into the collection at the specified index.
        body.getChildObjects().remove(paragraph);
        body.getChildObjects().insert(index, table);

        String result = "output/replaceTextWithTable.docx";

        // Save to file.
        document.saveToFile(result, FileFormat.Docx_2013);
    }
}
