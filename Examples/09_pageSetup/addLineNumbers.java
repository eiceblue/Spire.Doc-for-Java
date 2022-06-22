import com.spire.doc.*;

public class addLineNumbers {
    public static void main(String[] args) {
        //Create Word document.
        Document document = new Document();

        //Load the file from disk.
        document.loadFromFile("data/Template_Docx_1.docx");

        //Set the start value of the line numbers.
        document.getSections().get(0).getPageSetup().setLineNumberingStartValue(1);

        //Set the interval between displayed numbers.
        document.getSections().get(0).getPageSetup().setLineNumberingStep(6);

        //Set the distance between line numbers and text.
        document.getSections().get(0).getPageSetup().setLineNumberingDistanceFromText(40f);

        //Set the numbering mode of line numbers. There are four choices: None, Continuous, RestartPage and RestartSection.
        document.getSections().get(0).getPageSetup().setLineNumberingRestartMode(LineNumberingRestartMode.Continuous);

        String result = "output/result-addLineNumbers.docx";

        //Save to file.
        document.saveToFile(result, FileFormat.Docx_2013);
    }
}
