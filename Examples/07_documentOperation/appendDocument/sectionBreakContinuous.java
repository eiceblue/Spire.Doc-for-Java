import com.spire.doc.*;
import com.spire.doc.documents.*;

public class sectionBreakContinuous {
    public static void main(String[] args) {
        //Open a Word document
        Document doc = new Document();
        doc.loadFromFile("data/Sample_two sections.docx");

        for (Object sectionObj : doc.getSections()) {
            Section section=(Section)sectionObj;
            // Set section break as continuous
            section.setBreakCode(SectionBreakType.No_Break);
        }

        doc.saveToFile("output/sectionBreakContinuous.docx", FileFormat.Docx_2013);
    }
}
