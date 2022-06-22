import com.spire.doc.*;
import com.spire.doc.documents.*;

public class removeFooter {
    public static void main(String[] args) {
        String input = "data/oddAndEvenHeaderFooter.docx";
        String output = "output/removeFooter.docx";

        //load the document
        Document doc = new Document();
        doc.loadFromFile(input);

        //get the first section
        Section section = doc.getSections().get(0);

        //clear footer in the first page
        HeaderFooter footer;
        footer = section.getHeadersFooters().getByHeaderFooterType(HeaderFooterType.Footer_First_Page);
        if (footer != null)
            footer.getChildObjects().clear();
        //clear footer in the odd page
        footer = section.getHeadersFooters().getByHeaderFooterType(HeaderFooterType.Footer_Odd);
        if (footer != null)
            footer.getChildObjects().clear();
        //clear footer in the even page
        footer = section.getHeadersFooters().getByHeaderFooterType(HeaderFooterType.Footer_Even);
        if (footer != null)
            footer.getChildObjects().clear();

        //save the document
        doc.saveToFile(output, FileFormat.Docx);
    }
}
