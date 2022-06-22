import com.spire.doc.*;
import com.spire.doc.documents.*;

public class addPageNumbersInSections {
    public static void main(String[] args) {
        //Create Word document.
        Document document = new Document();

        //Load the file from disk.
        document.loadFromFile("data/Template_Docx_4.docx");

        //Repeat step2 and Step3 for the rest sections, so changing the code with for loop.
        for (int i = 0; i < 3; i++) {
            HeaderFooter footer = document.getSections().get(i).getHeadersFooters().getFooter();
            Paragraph footerParagraph = footer.addParagraph();
            footerParagraph.appendField("page number", FieldType.Field_Page);
            footerParagraph.appendText(" of ");
            footerParagraph.appendField("number of pages", FieldType.Field_Section_Pages);
            footerParagraph.getFormat().setHorizontalAlignment(HorizontalAlignment.Right);

            if (i == 2)
                break;
            else {
                document.getSections().get(i + 1).getPageSetup().setRestartPageNumbering(true);
                document.getSections().get(i + 1).getPageSetup().setPageStartingNumber(1);
            }
        }

        String result = "output/result-addPageNumbersInSections.docx";
        //Save to file.
        document.saveToFile(result, FileFormat.Docx_2013);
    }
}
