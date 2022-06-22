import com.spire.doc.*;
import com.spire.doc.documents.*;
import com.spire.doc.fields.*;

import java.awt.*;

public class insertEndnote {
    public static void main(String[] args) {
        //Create a document and load file
        Document doc = new Document();
        doc.loadFromFile( "data/insertEndnote.doc");
        Section s = doc.getSections().get(0);
        Paragraph p = s.getParagraphs().get(1);

        //Add endnote
        Footnote endnote = p.appendFootnote(FootnoteType.Endnote);

        //Append text
        TextRange text = endnote.getTextBody().addParagraph().appendText("Reference: Wikipedia");

        //Set text format
        text.getCharacterFormat().setFontName("Impact");
        text.getCharacterFormat().setFontSize(14);
        text.getCharacterFormat().setTextColor(new Color(255, 140, 0) /*Color.getDarkOrange()*/);

        //Set marker format of endnote
        endnote.getMarkerCharacterFormat().setFontName("Calibri");
        endnote.getMarkerCharacterFormat().setFontSize(20);
        endnote.getMarkerCharacterFormat().setTextColor(new Color(0, 0, 139)/*Color.getDarkBlue()*/);

        //Save the document
        String output = "output/insertEndnote.docx";
        doc.saveToFile(output, FileFormat.Docx);
    }
}
