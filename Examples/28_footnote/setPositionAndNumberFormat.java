import com.spire.doc.*;

public class setPositionAndNumberFormat {
    public static void main(String[] args) {
        //Load the document
        Document doc = new Document();
        doc.loadFromFile("data/footnote.docx");

        //Get the first section
        Section sec = doc.getSections().get(0);

        //Set the number format, restart rule and position for the footnote
        sec.getFootnoteOptions().setNumberFormat(FootnoteNumberFormat.Upper_Case_Letter);
        sec.getFootnoteOptions().setRestartRule(FootnoteRestartRule.Restart_Page);
        sec.getFootnoteOptions().setPosition(FootnotePosition.Print_As_End_Of_Section);

        //Save the document
        String output = "output/setPositionAndNumberFormat.docx";
        doc.saveToFile(output, FileFormat.Docx);
    }
}
