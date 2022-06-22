import com.spire.doc.*;
import com.spire.doc.documents.*;
import com.spire.doc.fields.*;
import java.awt.*;

public class imageHeaderAndFooter {
    public static void main(String[] args){
        String input1 = "data/multiPages.docx";
        String input2 = "data/e-iceblue.png";
        String input3 = "data/logo.png";
        String output = "output/imageHeaderAndFooter.docx";

        //load the document from disk
        Document doc = new Document();
        doc.loadFromFile(input1);

        //get the header of the first page
        HeaderFooter header = doc.getSections().get(0).getHeadersFooters().getHeader();

        //add a paragraph for the header
        Paragraph paragraph = header.addParagraph();

        //set the format of the paragraph
        paragraph.getFormat().setHorizontalAlignment(HorizontalAlignment.Right);

        //append a picture in the paragraph
        DocPicture headerimage = paragraph.appendPicture(input2);
        headerimage.setVerticalAlignment( ShapeVerticalAlignment.Bottom);

        //get the footer of the first section
        HeaderFooter footer = doc.getSections().get(0).getHeadersFooters().getFooter();

        //add a paragraph for the footer
        Paragraph paragraph2 = footer.addParagraph();

        //set the format of the paragraph
        paragraph2.getFormat().setHorizontalAlignment(HorizontalAlignment.Left);

        //append a picture in the paragraph
        DocPicture footerimage = paragraph2.appendPicture(input3);

        //append text in the paragraph
        TextRange TR = paragraph2.appendText("Copyright © 2013 e-iceblue. All Rights Reserved.");
        TR.getCharacterFormat().setFontName("Arial");
        TR.getCharacterFormat().setFontSize(10);
        TR.getCharacterFormat().setTextColor( Color.BLACK);

        //save the document
        doc.saveToFile(output, FileFormat.Docx);
    }
}
