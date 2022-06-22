import com.spire.doc.*;
import com.spire.doc.documents.*;
import com.spire.doc.fields.*;
import java.awt.*;

public class pageBorderSurround {
    public static void main(String[] args) {
        String output = "output/pageBorderSurround.docx";

        //create a new document
        Document doc = new Document();
        Section section = doc.addSection();

        //add a sample page border to the document
        section.getPageSetup().getBorders().setBorderType(BorderStyle.Wave);
        section.getPageSetup().getBorders().setColor(Color.GREEN);
        section.getPageSetup().getBorders().getLeft().setSpace( 20);
        section.getPageSetup().getBorders().getRight().setSpace( 20);

        //add a header and set its format
        Paragraph paragraph1 = section.getHeadersFooters().getHeader().addParagraph();
        paragraph1.getFormat().setHorizontalAlignment(HorizontalAlignment.Right);
        TextRange headerText = paragraph1.appendText("Header isn't included in page border");
        headerText.getCharacterFormat().setFontName("Calibri");
        headerText.getCharacterFormat().setFontSize( 20);
        headerText.getCharacterFormat().setBold(true);

        //add a footer and set its format
        Paragraph paragraph2 = section.getHeadersFooters().getFooter().addParagraph();
        paragraph2.getFormat().setHorizontalAlignment(HorizontalAlignment.Left);
        TextRange footerText = paragraph2.appendText("Footer is included in page border");
        footerText.getCharacterFormat().setFontName("Calibri");
        footerText.getCharacterFormat().setFontSize( 20);
        footerText.getCharacterFormat().setBold(true);

        //set the header not included in the page border while the footer included
        section.getPageSetup().setPageBorderIncludeHeader( false);
        section.getPageSetup().setHeaderDistance( 40);
        section.getPageSetup().setPageBorderIncludeFooter(true);
        section.getPageSetup().setFooterDistance( 40);

        //save the document
        doc.saveToFile(output, FileFormat.Docx);
    }
}
