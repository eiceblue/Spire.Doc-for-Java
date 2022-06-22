import com.spire.doc.*;

public class WPTtoDoc {
    public static void main(String[] args) {
        Document document = new Document();
        document.loadFromFile("data/Sample.wpt");
        document.saveToFile("output/WPTtoDoc.doc", FileFormat.Doc);
    }
}
