import com.spire.doc.*;
import com.spire.doc.documents.*;
import com.spire.doc.fields.*;

public class insertPageBreakFirstApproach {
    public static void main(String[] args) {
        //Create Word document.
        Document document = new Document();

        //Load the file from disk.
        document.loadFromFile("data/Template_Docx_2.docx");

        //Find the specified word "technology" where we want to insert the page break.
        TextSelection[] selections = document.findAllString("technology", false, true);

        //Traverse each word "technology".
        for(int i=0;i<selections.length;i++){
            TextSelection ts = selections[i];
            TextRange range = ts.getAsOneRange();
            Paragraph paragraph = range.getOwnerParagraph();
            int index = paragraph.getChildObjects().indexOf(range);

            //Create a new instance of page break and insert a page break after the word "technology".
            Break pageBreak = new Break(document, BreakType.Page_Break);
            paragraph.getChildObjects().insert(index + 1, pageBreak);
        }

        String result = "output/result-insertPageBreakAtSpecifiedLocation.docx";

        //Save to file.
        document.saveToFile(result, FileFormat.Docx_2013);
    }
}
