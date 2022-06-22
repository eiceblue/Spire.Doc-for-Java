import com.spire.doc.*;

public class addHeaderOnlyFirstPage {
    public static void main(String[] args) {
        String input1="data/headerSample.docx";
        String input2="data/multiPages.docx";
        String output="output/addHeaderOnlyFirstPage.docx";

        //load the source file
        Document doc1 = new Document();
        doc1.loadFromFile(input1);

        //get the header from the first section
        HeaderFooter header = doc1.getSections().get(0).getHeadersFooters().getHeader();

        //load the destination file
        Document doc2 = new Document();
        doc2.loadFromFile(input2);

        //get the first page header of the destination document
        HeaderFooter firstPageHeader = doc2.getSections().get(0).getHeadersFooters().getFirstPageHeader();

        //specify that the current section has a different header/footer for the first page
        for (int i=0;i<doc2.getSections().getCount();i++)
        {
            doc2.getSections().get(i).getPageSetup().setDifferentFirstPageHeaderFooter(true);
        }
        //removes all child objects in firstPageHeader
        firstPageHeader.getParagraphs().clear();;

        //add all child objects of the header to firstPageHeader
        for (int j=0; j< header.getChildObjects().getCount();j++)
        {
            DocumentObject obj = header.getChildObjects().get(j);
            firstPageHeader.getChildObjects().add(obj.deepClone());
        }
        //save the file
        doc2.saveToFile(output, FileFormat.Docx);
    }
}
