import com.spire.doc.*;
import com.spire.doc.documents.*;
import com.spire.doc.fields.*;

import java.awt.*;

public class insertFootnote {
    public static void main(String[] args) {
        //Create a document and load file
        Document document = new Document();
        document.loadFromFile("data/insertFootnote.docx");
        TextSelection selection = document.findString("Spire.Doc", false, true);

        //Add footnote
        TextRange textRange = selection.getAsOneRange();
        Paragraph paragraph = textRange.getOwnerParagraph();
        int index = paragraph.getChildObjects().indexOf(textRange);
        Footnote footnote = paragraph.appendFootnote(FootnoteType.Footnote);
        paragraph.getChildObjects().insert(index + 1, footnote);

        //Set text format
        textRange = footnote.getTextBody().addParagraph().appendText("Welcome to evaluate Spire.Doc");
        textRange.getCharacterFormat().setFontName("Arial Black");
        textRange.getCharacterFormat().setFontSize(10);
        textRange.getCharacterFormat().setTextColor(new Color(255, 140, 0));

        footnote.getMarkerCharacterFormat().setFontName("Calibri");
        footnote.getMarkerCharacterFormat().setFontSize(12);
        footnote.getMarkerCharacterFormat().setBold(true);
        footnote.getMarkerCharacterFormat().setTextColor(new Color(0, 0, 139));

        //Save the document
        String output = "output/insertFootnote.docx";
        document.saveToFile(output, FileFormat.Docx_2010);
    }
}
