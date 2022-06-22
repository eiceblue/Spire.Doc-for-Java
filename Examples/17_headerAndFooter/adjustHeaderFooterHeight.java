import com.spire.doc.*;

public class adjustHeaderFooterHeight {
    public static void main(String[] args) {
        String input="data/headerSample.docx";
        String output="output/adjustHeaderFooterHeight.docx";

        //load the document
        Document doc = new Document();
        doc.loadFromFile(input);

        //get the first section
        Section section = doc.getSections().get(0);

        //adjust the height of headers in the section
        section.getPageSetup().setHeaderDistance(100);

        //adjust the height of footers in the section
        section.getPageSetup().setFooterDistance(100);

        //save the document
        doc.saveToFile(output, FileFormat.Docx);
    }
}
