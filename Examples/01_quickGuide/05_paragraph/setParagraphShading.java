import com.spire.doc.*;
import com.spire.doc.documents.*;
import com.spire.doc.fields.*;

import java.awt.*;

public class setParagraphShading {
    public static void main(String[] args) {
        //Create Word document.
        Document document = new Document();

        //Load the file from disk.
        document.loadFromFile("data/Template_Docx_1.docx");

        //Get a paragraph.
        Paragraph paragaph = document.getSections().get(0).getParagraphs().get(0);

        //Set background color for the paragraph.
        paragaph.getFormat().setBackColor(Color.yellow);

        //Set background color for the selected text of paragraph.
        paragaph = document.getSections().get(0).getParagraphs().get(2);
        TextSelection selection = paragaph.find("Christmas", true, false);
        TextRange range = selection.getAsOneRange();
        range.getCharacterFormat().setTextBackgroundColor(Color.yellow);

        String result = "output/setParagraphShading.docx";

        //Save to file.
        document.saveToFile(result, FileFormat.Docx_2013);
    }
}
