import com.spire.doc.*;
import com.spire.doc.fields.DocPicture;

public class addCoverImage {
    public static void main(String[] args) {
        Document doc = new Document();
        doc.loadFromFile("data/ToEpub.doc");
        DocPicture picture = new DocPicture(doc);
        picture.loadImage("data/Cover.png");
        String result = "output/addCoverImage.epub";
        doc.saveToEpub(result, picture);
    }
}
