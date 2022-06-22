import com.spire.doc.*;
import com.spire.doc.documents.*;
import com.spire.doc.fields.TextRange;

import java.awt.*;

public class setHyperlinkFormat {
    public static void main(String[] args) {
        //Load a word document
        String input = "data/BlankTemplate.docx";
        Document doc = new Document();
        doc.loadFromFile(input);
        Section section = doc.getSections().get(0);

        //Add a paragraph and append a hyperlink to the paragraph
        Paragraph para1 = section.addParagraph();
        para1.appendText("Regular Link: ");
        //Format the hyperlink with default color and underline style
        TextRange txtRange1 = para1.appendHyperlink("www.e-iceblue.com", "www.e-iceblue.com", HyperlinkType.Web_Link);
        txtRange1.getCharacterFormat().setFontName("Times New Roman");
        txtRange1.getCharacterFormat().setFontSize(12f);
        Paragraph blankPara1 = section.addParagraph();

        //Add a paragraph and append a hyperlink to the paragraph
        Paragraph para2 = section.addParagraph();
        para2.appendText("Change Color: ");
        //Format the hyperlink with red color and underline style
        TextRange txtRange2 = para2.appendHyperlink("www.e-iceblue.com", "www.e-iceblue.com", HyperlinkType.Web_Link);
        txtRange2.getCharacterFormat().setFontName("Times New Roman");
        txtRange2.getCharacterFormat().setFontSize(12f);
        txtRange2.getCharacterFormat().setTextColor(Color.red);
        Paragraph blankPara2 = section.addParagraph();

        //Add a paragraph and append a hyperlink to the paragraph
        Paragraph para3 = section.addParagraph();
        para3.appendText("Remove Underline: ");
        //Format the hyperlink with red color and no underline style
        TextRange txtRange3 = para3.appendHyperlink("www.e-iceblue.com", "www.e-iceblue.com", HyperlinkType.Web_Link);
        txtRange3.getCharacterFormat().setFontName("Times New Roman");
        txtRange3.getCharacterFormat().setFontSize(12f);
        txtRange3.getCharacterFormat().setUnderlineStyle(UnderlineStyle.None);

        //Save the document
        String output = "output/HyperlinkFormat.docx";
        doc.saveToFile(output, FileFormat.Docx);
    }
}
