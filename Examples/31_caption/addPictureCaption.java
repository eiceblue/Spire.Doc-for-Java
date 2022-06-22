import com.spire.doc.*;
import com.spire.doc.documents.*;
import com.spire.doc.fields.*;

public class addPictureCaption {
    public static void main(String[] args) {
        //Create word document
        Document document = new Document();

        //Create a new section
        Section section = document.addSection();

        //Add the first picture
        Paragraph par1 = section.addParagraph();
        par1.getFormat().setAfterSpacing(10);
        DocPicture pic1 = par1.appendPicture("data/spire.Doc.png");
        pic1.setHeight(120);
        pic1.setWidth(120);
        //Add caption to the picture
        CaptionNumberingFormat format = CaptionNumberingFormat.Number;
        pic1.addCaption("Figure", format, CaptionPosition.Below_Item);

        //Add the second picture
        Paragraph par2 = section.addParagraph();
        DocPicture pic2 = par2.appendPicture("data/word.png");
        pic2.setHeight(120);
        pic2.setWidth(120);
        //Add caption to the picture
        pic2.addCaption("Figure", format, CaptionPosition.Below_Item);

        //Update fields
        document.isUpdateFields(true);

        //Save the file
        String output = "output/addPictureCaption_result.docx";
        document.saveToFile(output, FileFormat.Docx);

    }
}
