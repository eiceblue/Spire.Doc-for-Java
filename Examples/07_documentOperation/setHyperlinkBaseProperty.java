import com.spire.doc.*;

public class setHyperlinkBaseProperty {
    public static void main(String[] args) {
        String input = "data/setHyperlinkBaseProperty.docx";
        Document doc = new Document();
        doc.loadFromFile(input);
        //Set Hyperlink Base Property
        doc.getBuiltinDocumentProperties().setHyperLinkBase("HyperLinkBaseTest");
        //Save the document
        String output = "output/setHyperlinkBaseProperty_result.docx";
        doc.saveToFile(output, FileFormat.Docx);
    }
}
