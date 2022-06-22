import com.spire.doc.*;

public class copyHeaderAndFooter {
    public static void main(String[] args) {
        String input1="data/headerSample.docx";
        String input2="data/multiPages.docx";
        String output="output/copyHeaderAndFooter.docx";

        //load the source file
        Document doc1 = new Document();
        doc1.loadFromFile(input1);

        //get the header section from the source document
        HeaderFooter header = doc1.getSections().get(0).getHeadersFooters().getHeader();

        //load the destination file
        Document doc2 = new Document();
        doc2.loadFromFile(input2);

        //copy each object in the header of source file to destination file
        for (int i=0;i<doc2.getSections().getCount();i++)
        {
            for (int j=0; j< header.getChildObjects().getCount();j++)
            {
                DocumentObject obj = header.getChildObjects().get(j);
                doc2.getSections().get(i).getHeadersFooters().getHeader().getChildObjects().add(obj.deepClone());
            }
        }
        //save the document
        doc2.saveToFile(output, FileFormat.Docx);
    }
}
