import com.spire.doc.*;

public class xmlToPDF {
    public static void main(String[] args) {

        String inputFile="data/Template_XmlFile.xml";
        String outputFile="output/Result-XMLToPDF.pdf";

        Document document = new Document();
        document.loadFromFile(inputFile);
        document.saveToFile(outputFile, FileFormat.PDF );
    }
}
