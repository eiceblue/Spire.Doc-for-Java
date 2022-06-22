import com.spire.doc.*;

public class toPCL {
    public static void main(String[] args) {
        String input = "data/convertedTemplate.docx";
        String output = "output/toPCL.PCL";

        //load Word document
        Document document= new Document();
        document.loadFromFile(input);

        //save the document to PCL format
        document.saveToFile(output, FileFormat.PCL);
    }
}
