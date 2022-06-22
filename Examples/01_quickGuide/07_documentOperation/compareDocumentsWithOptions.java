import com.spire.doc.*;
import com.spire.doc.documents.comparison.*;

public class compareDocumentsWithOptions {
    public static void main(String[] args){
        //Load the first document
        Document doc1 = new Document();
        doc1.loadFromFile("data/SupportDocumentCompare1.docx");
        //Load the second document
        Document doc2 = new Document();
        doc2.loadFromFile("data/SupportDocumentCompare2.docx");

        CompareOptions compareOptions = new CompareOptions();
        compareOptions.setIgnoreFormatting(true);
        //Compare the two documents
        doc1.compare(doc2, "E-iceblue",compareOptions);

        String result = "output/CompareDocumentsWithOptions_result.docx";
        //Save the comparison document
        doc1.saveToFile(result, FileFormat.Docx_2013);
    }
}
