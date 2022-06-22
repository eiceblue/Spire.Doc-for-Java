import com.spire.doc.*;

public class DocToWPT {
    public static void main(String[] args) {
        Document document = new Document();
        document.loadFromFile("data/Sample.doc");
        document.saveToFile("output/DocToWPT.wpt", FileFormat.WPT);
    }
}
