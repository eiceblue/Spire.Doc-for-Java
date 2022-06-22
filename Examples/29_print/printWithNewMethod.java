import com.spire.doc.*;

import java.awt.print.*;

public class printWithNewMethod {
    public static void main(String[] args) {
        Document document = new Document();
        document.loadFromFile("data/print.docx");
        PrinterJob loPrinterJob = PrinterJob.getPrinterJob();
        PageFormat loPageFormat  = loPrinterJob.defaultPage();
        Paper loPaper = loPageFormat.getPaper();
        //delete the default margin
        loPaper.setImageableArea(0,0,loPageFormat.getWidth(),loPageFormat.getHeight());
        //set the copy number
        loPrinterJob.setCopies(1);
        loPageFormat.setPaper(loPaper);
        loPrinterJob.setPrintable(document,loPageFormat);
        try {
            loPrinterJob.print();
        } catch (PrinterException e)
        {
            e.printStackTrace();
        }

    }
}
