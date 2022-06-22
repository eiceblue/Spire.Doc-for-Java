import com.spire.doc.*;

public class htmlToXml {
    public static void main(String[] args) {
        //Create Word document.
        Document document = new Document();

        //Load the file from disk.
        document.loadFromFile("data/Template_HtmlFile.html");

        String result = "output/result-HtmlToXml.xml";

        //Save to file.
        document.saveToFile(result, FileFormat.Xml);
    }
}
