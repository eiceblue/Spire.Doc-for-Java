import com.spire.doc.Document;
import com.spire.doc.FileFormat;

public class wordtoWordXML {
    public static void main(String[] args) {

        String inputFile="data/Template_Docx_1.docx";
        String result1 = "output/Result-WordToWordML.xml";
        String result2 = "output/Result-WordToWordXML.xml";

        Document document = new Document();
        document.loadFromFile(inputFile);

        //For Word 2003:
        document.saveToFile(result1, FileFormat.Word_ML);

        //For Word 2007:
        document.saveToFile(result2, FileFormat.Word_Xml);
    }
}
