import com.spire.doc.*;
import com.spire.doc.documents.*;
import java.awt.*;

public class addPageBorders {
    public static void main(String[] args) {
        //Create Word document.
        Document document = new Document();

        //Load the file from disk.
        document.loadFromFile("data/Template_Docx_1.docx");

        //Get the first section.
        Section section = document.getSections().get(0);

        //Add page borders with special style and color.
        section.getPageSetup().getBorders().setBorderType(BorderStyle.Double_Wave);
        section.getPageSetup().getBorders().setColor(Color.lightGray);

        //Set the space between border and text.
        section.getPageSetup().getBorders().getLeft().setSpace(50);
        section.getPageSetup().getBorders().getRight().setSpace(50);

        String result = "output/result-addPageBorders.docx";

        //Save to file.
        document.saveToFile(result, FileFormat.Docx_2013);
    }

}