import com.spire.doc.*;

public class rtfToPDF {
    public static void main(String[] args) {
        //Create Word document.
        Document document = new Document();

        //Load the file from disk.
        document.loadFromFile( "data/Template_RtfFile.rtf", FileFormat.Rtf);

        String result = "output/result-rtfToPdf.pdf";

        //Save to file.
        document.saveToFile(result, FileFormat.PDF);
    }
}
