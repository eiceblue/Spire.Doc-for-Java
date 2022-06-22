import com.spire.doc.*;
import com.spire.doc.documents.*;
import com.spire.doc.fields.*;

public class setSuperscriptAndSubscript {
    public static void main(String[] args) {
        //Create word document
        Document document = new Document();

        //Create a new section
        Section section = document.addSection();

        Paragraph paragraph = section.addParagraph();
        paragraph.appendText("E = mc");
        TextRange range1 = paragraph.appendText("2");

        //Set superscript
        range1.getCharacterFormat().setSubSuperScript(SubSuperScript.Super_Script);

        paragraph.appendBreak(BreakType.Line_Break);
        paragraph.appendText("F");
        TextRange range2 = paragraph.appendText("n");

        //Set subscript
        range2.getCharacterFormat().setSubSuperScript(SubSuperScript.Sub_Script);

        paragraph.appendText(" = F");
        paragraph.appendText("n-1").getCharacterFormat().setSubSuperScript(SubSuperScript.Sub_Script);
        paragraph.appendText(" + F");
        paragraph.appendText("n-2").getCharacterFormat().setSubSuperScript(SubSuperScript.Sub_Script);

        //Set font size
        for (Object i : paragraph.getItems()) {
            if (i instanceof TextRange) {
                ((TextRange) i).getCharacterFormat().setFontSize(36);
            }
        }

        //Save the file
        String output = "output/setSuperscriptAndSubscript.docx";
        document.saveToFile(output, FileFormat.Docx);
    }
}
