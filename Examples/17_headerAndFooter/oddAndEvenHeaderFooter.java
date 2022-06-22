import com.spire.doc.*;
import com.spire.doc.documents.*;
import com.spire.doc.fields.*;

public class oddAndEvenHeaderFooter {
    public static void main(String[] args) {
        String input = "data/multiPages.docx";
        String output = "output/oddAndEvenHeaderFooter.docx";

        //load the document
        Document doc = new Document();
        doc.loadFromFile(input);

        //get the section and
        Section section = doc.getSections().get(0);

        //set the DifferentOddAndEvenPagesHeaderFooter property to ture
        section.getPageSetup().setDifferentOddAndEvenPagesHeaderFooter(true);

        //add odd header
        Paragraph P3 = section.getHeadersFooters().getOddHeader().addParagraph();
        TextRange OH = P3.appendText("Odd Header");
        P3.getFormat().setHorizontalAlignment(HorizontalAlignment.Center);
        OH.getCharacterFormat().setFontName("Arial");
        OH.getCharacterFormat().setFontSize( 10);

        //add even header
        Paragraph P4 = section.getHeadersFooters().getEvenHeader().addParagraph();
        TextRange EH = P4.appendText("Even Header from E-iceblue Using Spire.Doc");
        P4.getFormat().setHorizontalAlignment(HorizontalAlignment.Center);
        EH.getCharacterFormat().setFontName("Arial");
        EH.getCharacterFormat().setFontSize( 10);

        //add odd footer
        Paragraph P2 = section.getHeadersFooters().getOddFooter().addParagraph();
        TextRange OF = P2.appendText("Odd Footer");
        P2.getFormat().setHorizontalAlignment(HorizontalAlignment.Center);
        OF.getCharacterFormat().setFontName("Arial");
        OF.getCharacterFormat().setFontSize( 10);

        //add even footer
        Paragraph P1 = section.getHeadersFooters().getEvenFooter().addParagraph();
        TextRange EF = P1.appendText("Even Footer from E-iceblue Using Spire.Doc");
        EF.getCharacterFormat().setFontName("Arial");
        EF.getCharacterFormat().setFontSize( 10);
        P1.getFormat().setHorizontalAlignment(HorizontalAlignment.Center);

        //save the document
        doc.saveToFile(output, FileFormat.Docx);
    }
}
