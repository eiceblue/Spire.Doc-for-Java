import com.spire.doc.*;
import com.spire.ms.Printing.*;

public class customPaperSize {
    public static void main(String[] args) {
        //Load document
        Document doc = new Document();
        doc.loadFromFile("data/print.docx");

        //Get the PrintDocument object
        PrintDocument printDoc = doc.getPrintDocument();

        //Custom the paper size
        PaperSize size = new PaperSize();
        size.setWidth(900);
        size.setHeight(800);
        printDoc.getDefaultPageSettings().setPaperSize(size);

        //Print the document
        printDoc.print();
    }
}
