import com.spire.doc.*;

public class disableHyperlinks {
    public static void main(String[] args) {

        String inputFile="data/Template_Docx_5-BJ.docx";
        String outputFile="output/Result-DisableHyperlinks.pdf";

        //create Word document.
        Document document = new Document();

        //load the file from disk.
        document.loadFromFile(inputFile);

        //create an instance of ToPdfParameterList.
        ToPdfParameterList pdf = new ToPdfParameterList();

        //set setDisableLink to true to remove the hyperlink effect for the result PDF page.
        //set setDisableLink to false to preserve the hyperlink effect for the result PDF page.
        pdf.setDisableLink(true);

        //save to file.
        document.saveToFile(outputFile,pdf);
    }
}
