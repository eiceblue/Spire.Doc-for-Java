import com.spire.doc.*;

public class xmlToWord {
    public static void main(String[] args) {

        String inputFile="data/Template_XmlFile.xml";
        String outputFile="output/Result-XMLToWord.docx";

        Document document = new Document();
        document.loadFromFile(inputFile);
        document.saveToFile(outputFile, FileFormat.Docx );
    }
}
