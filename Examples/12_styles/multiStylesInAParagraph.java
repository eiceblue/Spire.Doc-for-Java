import com.spire.doc.*;
import com.spire.doc.documents.*;
import com.spire.doc.fields.TextRange;

import java.awt.*;

public class multiStylesInAParagraph {
    public static void main(String[] args) {

        String outputFile="output/multiStylesInAParagraph.docx";

        //Create a Word document
        Document doc = new Document();

        //Add a section
        Section section = doc.addSection();

        //Add a paragraph
        Paragraph para = section.addParagraph();

        //Add a text range 1 and set its style
        TextRange range = para.appendText("Spire.Doc for Java ");
        range.getCharacterFormat().setFontName("Calibri");
        range.getCharacterFormat().setFontSize(16);
        range.getCharacterFormat().setTextColor(Color.blue);
        range.getCharacterFormat().setBold(true);
        range.getCharacterFormat().setUnderlineStyle(UnderlineStyle.Single);

        //Add a text range 2 and set its style
        range = para.appendText("is a professional Word Java library");
        range.getCharacterFormat().setFontName("Calibri");
        range.getCharacterFormat().setFontSize(15);

        //Save the Word document
        doc.saveToFile(outputFile, FileFormat.Docx);
    }
}
