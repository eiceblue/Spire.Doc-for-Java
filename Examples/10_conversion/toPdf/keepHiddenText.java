import com.spire.doc.*;

public class keepHiddenText {
    public static void main(String[] args) {

        String inputFile="data/Template_Docx_5-BJ.docx";
        String outputFile="output/Result-SaveTheHiddenTextToPDF.pdf";

        //create Word document.
        Document document = new Document();

        //load the file from disk.
        document.loadFromFile(inputFile);

        //when converting to PDF file, set the property isHidden as true.
        ToPdfParameterList pdf = new ToPdfParameterList();
        pdf.isHidden(true);

        //save to file.
        document.saveToFile(outputFile,pdf);
    }
}
