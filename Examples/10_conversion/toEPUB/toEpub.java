import com.spire.doc.*;

public class toEpub {
    public static void main(String[] args) {
        Document doc = new Document();
        doc.loadFromFile("data/ToEpub.doc");
        String result = "output/toEpub.epub";
        doc.saveToFile(result,FileFormat.E_Pub);
    }
}
