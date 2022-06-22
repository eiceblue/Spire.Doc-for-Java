import com.spire.doc.Document;

/**
 * Convert doc/docx to tiff.
 */
public class toTiff {
    public static void main(String[] args) {
        String outputFile="output/wordToTiff.tiff";
        String sampleText = "\nSpire.Doc for Java is a professional Word API that empowers Java applications to create, convert, " +
                "manipulate and print Word documents without dependency on Microsoft Word.\n\nBy using this multifunctional library," +
                " developers are able to process copious tasks effortlessly, such as inserting image, hyperlink, digital signature," +
                " bookmark and watermark, setting header and footer, creating table, setting background image, and adding footnote" +
                " and endnote.\n\nIn addition, Spire.Doc for Java supports file format conversions from Word to PDF, XPS, Image, EPUB" +
                ", HTML, TXT, ODT, RTF, WordML, WordXML and many more.";
        Document doc = new Document();
        //Append some text to the document
        doc.addSection().addParagraph().appendText(sampleText);
        //Convert Word to Tiff
        doc.saveToTiff(outputFile);
        doc.close();
    }
}
