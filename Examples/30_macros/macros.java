import com.spire.doc.*;

public class macros {
    public static void main(String[] args) {
        //Load Word document which contains macro
        String input = "data/macros.docm";
        Document doc = new Document();
        doc.loadFromFile(input, FileFormat.Docm);

        //Save the file
        String output = "output/macros.docm";
        doc.saveToFile(output,FileFormat.Docm);
    }
}
