import com.spire.doc.*;

public class addAlternativeText {
    public static void main(String[] args) {
        String input = "data/tableSample.docx";
        String output = "output/addAlternativeText.docx";

        //load the document
        Document doc = new Document();
        doc.loadFromFile(input);

        //get the first section
        Section section = doc.getSections().get(0);

        //get the first table in the section
        Table table = (Table)section.getTables().get(0);

        //add title
        table.setTitle("Table 1");

        //add description
        table.setTableDescription("Description Text");

        //save the document
        doc.saveToFile(output, FileFormat.Docx);
    }
}
