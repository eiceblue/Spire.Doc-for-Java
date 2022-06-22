import com.spire.doc.*;

public class setImageQuality {
    public static void main(String[] args) {

        String inputFile="data/Template_Doc_1.doc";
        String outputFile="output/Result-DocToPDFImageQuality.pdf";

        //Create Word document.
        Document document = new Document();

        //Load the file from disk.
        document.loadFromFile(inputFile);

        //Set the output image quality to be 40% of the original image. The default set of the output image quality is 80% of the original.
        document.setJPEGQuality(40);

        //Save to file.
        document.saveToFile(outputFile, FileFormat.PDF);

    }
}
