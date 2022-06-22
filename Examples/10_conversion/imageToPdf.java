import com.spire.doc.*;
import com.spire.doc.documents.*;
import com.spire.doc.fields.DocPicture;

public class imageToPdf {
    public static void main(String[] args) {
        String input = "data/Image.png";
        //Create a new document
        Document doc = new Document();
        //Create a new section
        Section section = doc.addSection();
        //Create a new paragraph
        Paragraph paragraph = section.addParagraph();
        //Add a picture for paragraph
        DocPicture picture = paragraph.appendPicture(input);
        //Set the page size to the same size as picture
        //section.PageSetup.PageSize = new SizeF(picture.Width, picture.Height);
        //Set A4 page size
        section.getPageSetup().setPageSize(PageSize.A4);
        //Set the page margins
        section.getPageSetup().getMargins().setTop(10f);
        section.getPageSetup().getMargins().setLeft(25f);

        String result = "output/imageToPdf.pdf";
        doc.saveToFile(result, FileFormat.PDF);
    }
}
