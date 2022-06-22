import com.spire.doc.*;

public class toPostScript {
    public static void main(String[] args) {
        String input = "data/convertedTemplate.docx";
        String output = "output/Result-toPostScript.ps";

        Document document= new Document();
        document.loadFromFile(input);
        document.saveToFile(output, FileFormat.Post_Script);
    }
}
