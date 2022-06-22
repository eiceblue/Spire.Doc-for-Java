import com.spire.doc.*;
import com.spire.doc.documents.*;
import com.spire.doc.fields.*;

public class hideParagraph {
    public static void main(String[] args) {
        //Create Word document.
        Document document = new Document();

        // Load the file from disk.
        document.loadFromFile("data/Template_Docx_1.docx");

        //Get the first section and the first paragraph from the word document.
        Section sec = document.getSections().get(0);
        Paragraph para = sec.getParagraphs().get(0);

        // Loop through the textranges and set CharacterFormat.Hidden property as true to hide the texts.
        for (Object docObj : para.getChildObjects()) {
            DocumentObject obj = (DocumentObject)docObj;
            if ((obj instanceof TextRange)) {
                TextRange range = ((TextRange)(obj));
                range.getCharacterFormat().setHidden(true);
            }
        }

        String result = "output/hideParagraph.docx";

        // Save to file.
        document.saveToFile(result, FileFormat.Docx_2013);
    }
}
