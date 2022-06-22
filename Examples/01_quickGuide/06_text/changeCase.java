import com.spire.doc.*;
import com.spire.doc.documents.*;
import com.spire.doc.fields.*;

public class changeCase {
    public static void main(String[] args) {
        // Create a new document and load from file
        String input = "data/Text1.docx"; ;
        Document doc = new Document();
        doc.loadFromFile(input);

        TextRange textRange;
        //Get the first paragraph and set its CharacterFormat to AllCaps
        Paragraph para1 = doc.getSections().get(0).getParagraphs().get(1);
        for (Object docObj : para1.getChildObjects()) {
            DocumentObject obj=(DocumentObject)docObj;
            if ((obj instanceof TextRange)) {
                textRange = ((TextRange)(obj));
                textRange.getCharacterFormat().setAllCaps(true);
            }
        }

        //Get the third paragraph and set its CharacterFormat to IsSmallCaps
        Paragraph para2 = doc.getSections().get(0).getParagraphs().get(3);
        for (Object docObj : para2.getChildObjects()) {
            DocumentObject obj=(DocumentObject)docObj;
            if ((obj instanceof TextRange)) {
                textRange = ((TextRange)(obj));
                textRange.getCharacterFormat().isSmallCaps(true);
            }
        }

        //Save and launch the document
        String output = "output/changeCase.docx";
        doc.saveToFile(output, FileFormat.Docx_2013);
    }
}
