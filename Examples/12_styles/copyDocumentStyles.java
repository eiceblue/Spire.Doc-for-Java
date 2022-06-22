import com.spire.doc.*;
import com.spire.doc.collections.StyleCollection;

public class copyDocumentStyles {
    public static void main(String[] args) {

        String inputFile1 ="data/Template_Toc.docx";
        String inputFile2 ="data/Template_Docx_1.docx";
        String outputFile="output/copyDocumentStyles.docx";

        //Load source document from disk
        Document srcDoc = new Document();
        srcDoc.loadFromFile(inputFile1);

        //Load destination document from disk
        Document destDoc= new Document();
        destDoc.loadFromFile(inputFile2);

        //Get the style collections of source document
        StyleCollection styles = srcDoc.getStyles();

        //Add the style to destination document
        for (int i = 0; i < styles.getCount(); i ++)
        {
            destDoc.getStyles().add(styles.get(i));
        }

        //Save the Word file
        destDoc.saveToFile(outputFile, FileFormat.Docx);
    }
}
