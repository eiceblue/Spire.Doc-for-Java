import com.spire.doc.*;
import com.spire.doc.documents.*;
import com.spire.doc.fields.DocPicture;

public class createImageHyperlink {
    public static void main(String[] args) {
        String input = "data/BlankTemplate.docx";
        Document doc = new Document();
        doc.loadFromFile(input);

        Section section = doc.getSections().get(0);
        //Add a paragraph
        Paragraph paragraph = section.addParagraph();
        //Load an image to a DocPicture object
        DocPicture picture = new DocPicture(doc);
        picture.loadImage("data/SpireDocJava.png");
        //Add an image hyperlink to the paragraph
        paragraph.appendHyperlink("https://www.e-iceblue.com/Introduce/word-for-net-introduce.html", picture, HyperlinkType.Web_Link);

        //Save and launch document
        String output = "output/CreateImageHyperlink.docx";
        doc.saveToFile(output, FileFormat.Docx);
    }
}
