import com.spire.doc.*;

public class compareDocuments {
    public static void main(String[] args) {
        //load the first document
        Document doc1 = new Document();
        doc1.loadFromFile("data/DocumentCompareFunction1.docx");
        //load the second document
        Document doc2 = new Document();
        doc2.loadFromFile("data/DocumentCompareFunction2.docx");
        //compare two documents
        doc1.compare(doc2, "E-iceblue");
        //Save document
        String output = "output/CompareDocuments_result.docx";
        doc1.saveToFile(output, FileFormat.Docx);
    }
}
