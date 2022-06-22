import com.spire.doc.*;
import com.spire.doc.documents.*;
import java.awt.*;

public class findAndHighlight {
    public static void main(String[] args) {
        //Create word document.
        Document document = new Document();

        // Load the document from disk.
        document.loadFromFile("data/Sample.docx");

        //Find text.
        TextSelection[] textSelections = document.findAllString("word", false, true);

        // Set highlight.
        for (TextSelection selection : textSelections){
            selection.getAsOneRange().getCharacterFormat().setHighlightColor(Color.yellow);
        }

        // Save to file.
        document.saveToFile("output/findAndHighlight.docx", FileFormat.Docx_2013);
    }
}
