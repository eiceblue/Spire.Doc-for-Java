import com.spire.doc.*;
import com.spire.doc.documents.*;
import com.spire.doc.fields.*;

public class differentFirstPage {
    public static void main(String[] args) {
        String input1="data/multiPages.docx";
        String input2="data/e-iceblue.png";
        String output="output/differentFirstPage.docx";

        //load the document
        Document doc = new Document();
        doc.loadFromFile(input1);

        //get the section and set the property true
        Section section = doc.getSections().get(0);
        section.getPageSetup().setDifferentFirstPageHeaderFooter(true);

        //set the first page header. Here we append a picture in the header
        Paragraph paragraph1 = section.getHeadersFooters().getFirstPageHeader().addParagraph();
        paragraph1.getFormat().setHorizontalAlignment(HorizontalAlignment.Right);
        paragraph1.appendPicture(input2);

        //set the first page footer
        Paragraph paragraph2 = section.getHeadersFooters().getFirstPageFooter().addParagraph();
        paragraph2.getFormat().setHorizontalAlignment(HorizontalAlignment.Center);
        TextRange FF = paragraph2.appendText("First Page Footer");
        FF.getCharacterFormat().setFontSize( 10);

        //set the other header & footer.
        Paragraph paragraph3 = section.getHeadersFooters().getHeader().addParagraph();
        paragraph3.getFormat().setHorizontalAlignment(HorizontalAlignment.Center);
        TextRange NH = paragraph3.appendText("Spire.Doc for JAVA");
        NH.getCharacterFormat().setFontSize( 10);

        Paragraph paragraph4 = section.getHeadersFooters().getFooter().addParagraph();
        paragraph4.getFormat().setHorizontalAlignment(HorizontalAlignment.Center);
        TextRange NF = paragraph4.appendText("E-iceblue");
        NF.getCharacterFormat().setFontSize( 10);

        //save and launch document
        doc.saveToFile(output, FileFormat.Docx);
    }
}
