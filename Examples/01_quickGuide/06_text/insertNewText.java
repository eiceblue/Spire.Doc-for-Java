import com.spire.doc.*;
import com.spire.doc.documents.*;
import com.spire.doc.fields.*;
import java.awt.*;

public class insertNewText {
    public static void main(String[] args) {
        //Load Document
        String input = "data/Sample.docx";
        Document doc = new Document();
        doc.loadFromFile(input);

        //Find all the text string “New Zealand” from the sample document
        TextSelection[] selections = doc.findAllString("Word", true, true);
        int index = 0;

        //Defines text range
        TextRange range = new TextRange(doc);

        //Insert new text string (New) after the searched text string
        for (TextSelection selection : selections) {
            range = selection.getAsOneRange();
            TextRange newrange = new TextRange(doc);
            newrange.setText("(New text)");
            index = range.getOwnerParagraph().getChildObjects().indexOf(range);
            range.getOwnerParagraph().getChildObjects().insert((index + 1), newrange);
        }

        //Find and highlight the newly added text string New
        TextSelection[] text = doc.findAllString("New text", true, true);
        for (TextSelection seletion : text) {
            seletion.getAsOneRange().getCharacterFormat().setHighlightColor(Color.yellow);
        }

        //Save and launch document
        String output = "output/insertNewText.docx";
        doc.saveToFile(output, FileFormat.Docx);
    }
}
