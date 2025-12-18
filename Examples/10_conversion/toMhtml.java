import com.spire.doc.Document;
import com.spire.doc.FileFormat;

public class toMhtml {
    public static void main(String[] args) {

        //Create word document
        Document document = new Document();

        //Load the file from disk.
        document.loadFromFile("data\\ToMhtml.docx");

        //Save to RTF file.
        document.saveToFile("ToMhtml-out.mhtml", FileFormat.Mhtml);

        //Dispose the document
        document.dispose();

    }
}
