import com.spire.doc.*;
import com.spire.doc.documents.*;

public class htmlToPdf {
    public static void main(String[] args) {
        //Create Word document.
        Document document = new Document();

        //Load the file from disk.
        document.loadFromFile("data/Template_HtmlFile.html", FileFormat.Html, XHTMLValidationType.None);

        String result = "output/result-HtmlToPdf.pdf";

        //Save to file.
        document.saveToFile(result, FileFormat.PDF);
    }
}
