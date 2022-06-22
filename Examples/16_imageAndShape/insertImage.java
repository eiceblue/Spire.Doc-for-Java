import com.spire.doc.*;
import com.spire.doc.documents.*;
import com.spire.doc.fields.*;

public class insertImage {
    public static void main(String[] args){
        String input1 = "data/blankTemplate.docx";
        String input2 = "data/spireDoc.png";
        String output = "output/insertImage.docx";

        //load a document
        Document doc = new Document();
        doc.loadFromFile(input1);

        //add paragraphs
        Section section = doc.getSections().get(0);
        Paragraph paragraph = section.getParagraphs().getCount()> 0 ? section.getParagraphs().get(0): section.addParagraph();
        paragraph.appendText("The sample demonstrates how to insert an image into a document.");
        paragraph.applyStyle(BuiltinStyle.Heading_2);
        paragraph = section.addParagraph();
        paragraph.appendText("This is an inserted picture.");

        //add picture
        DocPicture picture = new DocPicture(doc);
        picture.loadImage(input2);

        //set image's position
        picture.setHorizontalPosition(50.0F);
        picture.setVerticalPosition(60.0F);

        //set image's size
        picture.setWidth(200);
        picture.setHeight(200);

        //set textWrappingStyle with image;
        picture.setTextWrappingStyle( TextWrappingStyle.Through);

        //insert the picture at the beginning of added second paragraph
        paragraph.getChildObjects().insert(0,picture);

        //save the document
        doc.saveToFile(output, FileFormat.Docx);
    }
}
