import com.spire.doc.*;

public class txtToWord {
    public static void main(String[] args) {
        String input = "data/Template_TxtFile.txt";
        String output = "output/Result-TxtToWord.docx";


        Document document= new Document();
        document.loadFromFile(input);
        document.saveToFile(output, FileFormat.Docx);
    }
}
