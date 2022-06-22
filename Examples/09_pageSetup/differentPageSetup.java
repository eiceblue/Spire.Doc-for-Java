import com.spire.doc.*;
import com.spire.doc.documents.*;

public class differentPageSetup {
    public static void main(String[] args) {
        //Open a Word document
        Document doc = new Document("data/DifferentPageSetup.docx");

        //Get the second section
        Section sectionOne = doc.getSections().get(0);

        //Set the orientation
        sectionOne.getPageSetup().setOrientation(PageOrientation.Landscape);

        //Set page size
        Section sectionTwo = doc.getSections().get(1);
        sectionTwo.getPageSetup().setPageSize(new Dimension(800, 800));

        String result = "output/result-differentPageSetup.docx";
        doc.saveToFile(result,FileFormat.Docx_2013);
    }
}
