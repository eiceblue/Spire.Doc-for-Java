import com.spire.doc.*;

public class wordToTxt {
    public static void main(String[] args) {
        String input = "data/convertedTemplate.docx";
        String output = "output/Result-WordToText.txt";

        Document document= new Document();
        document.loadFromFile(input);
        document.saveToFile(output, FileFormat.Txt);
    }
}
