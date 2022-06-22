import com.spire.doc.*;
import com.spire.doc.documents.*;

public class removeHeader {
    public static void main(String[] args) {
        String input = "data/oddAndEvenHeaderFooter.docx";
        String output = "output/removeHeader.docx";

        //load the document
        Document doc = new Document();
        doc.loadFromFile(input);

        //get the first section of the document
        Section section = doc.getSections().get(0);

        //clear footer in the first page
        HeaderFooter header;
        header = section.getHeadersFooters().getByHeaderFooterType(HeaderFooterType.Header_First_Page);
        if (header != null)
            header.getChildObjects().clear();
        //clear footer in the odd page
        header = section.getHeadersFooters().getByHeaderFooterType(HeaderFooterType.Header_Odd);
        if (header != null)
            header.getChildObjects().clear();
        //clear footer in the even page
        header = section.getHeadersFooters().getByHeaderFooterType(HeaderFooterType.Header_Even);
        if (header != null)
            header.getChildObjects().clear();

        //save the document
        doc.saveToFile(output, FileFormat.Docx);
    }
}
