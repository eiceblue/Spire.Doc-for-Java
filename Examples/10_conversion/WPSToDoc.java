import com.spire.doc.*;

public class WPSToDoc {
    public static void main(String[] args) {
        Document document = new Document();
        document.loadFromFile("data/Sample.wps");
        document.saveToFile("output/WPSToDoc.doc", FileFormat.Doc);
    }
}
