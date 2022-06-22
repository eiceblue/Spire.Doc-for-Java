import com.spire.doc.*;

public class removeTable {
    public static void main(String[] args) {
        //Load the document
        String input = "data/JAVATemplate_N.docx";
        Document doc = new Document();
        doc.loadFromFile(input);

        //Remove the first Table
        doc.getSections().get(0).getTables().removeAt(0);

        //Save and launch document
        String output = "output/RemoveTable.docx";
        doc.saveToFile(output, FileFormat.Docx);
    }
}
