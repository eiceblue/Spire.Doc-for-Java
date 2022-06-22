import com.spire.doc.*;

public class DocToWPS {
    public static void main(String[] args) {
        Document document = new Document();
        document.loadFromFile("data/Sample.doc");
        document.saveToFile("output/DocToWPS.wps", FileFormat.WPS);
    }
}
