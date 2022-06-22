import com.spire.doc.*;
import com.spire.doc.documents.*;

public class insertSectionBreak {
    public static void main(String[] args) {
        //Create Word document.
        Document document = new Document();

        //Load the file from disk.
        document.loadFromFile("data/Template_Docx_1.docx");

        //Insert section break. There are five section break options including EvenPage, NewColumn, NewPage, NoBreak, OddPage.
        document.getSections().get(0).getParagraphs().get(1).insertSectionBreak(SectionBreakType.No_Break);

        String result = "output/result-insertSectionBreak.docx";

        //Save the file.
        document.saveToFile(result, FileFormat.Docx_2013);
    }
}
