import com.spire.doc.*;
import com.spire.doc.documents.*;

public class insertPageBreakSecondApproach {
    public static void main(String[] args) {
        //Create Word document.
        Document document = new Document();

        //Load the file from disk.
        document.loadFromFile( "data/Template_Docx_1.docx");

        //Insert page break.
        document.getSections().get(0).getParagraphs().get(3).appendBreak(BreakType.Page_Break);

        String result = "output/result-insertPageBreakSecondApproach.docx";

        //Save the file.
        document.saveToFile(result, FileFormat.Docx_2013);
    }
}
