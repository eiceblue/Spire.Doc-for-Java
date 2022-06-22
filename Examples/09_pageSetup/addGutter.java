import com.spire.doc.*;
import java.io.IOException;

public class addGutter {
    public static void main(String[] args) throws IOException {
        String input = "data/addGutter.docx";
        String output = "output/addGutter_output.docx";

        //create word document
        Document document = new Document();

        //load the document from disk
        document.loadFromFile(input);

        //get the first section
        Section section = document.getSections().get(0);

        //set gutter
        section.getPageSetup().setGutter(100f);

        //save the file
        document.saveToFile(output, FileFormat.Docx);
    }
}
