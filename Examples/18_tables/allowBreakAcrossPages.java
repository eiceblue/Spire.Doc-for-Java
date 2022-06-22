import com.spire.doc.*;

public class allowBreakAcrossPages {
    public static void main(String[] args) {
        //Create word document
        Document document = new Document();
        document.loadFromFile("data/AllowBreakAcrossPages.docx");

        //Get the first section
        Section section = document.getSections().get(0);

        //Get the first table of the section
        Table table = section.getTables().get(0);

        for (int i=0;i<table.getRows().getCount();i++)
        {
            TableRow row=table.getRows().get(i);
            //Allow break across pages
            row.getRowFormat().isBreakAcrossPages(true);
        }

        //Save the Word document
        String output = "output/AllowBreakAcrossPages_out.docx";
        document.saveToFile(output, FileFormat.Docx_2013);
    }
}
