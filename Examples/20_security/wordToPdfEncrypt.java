import com.spire.doc.*;
import com.spire.pdf.security.*;

public class wordToPdfEncrypt {
    public static void main(String[] args) {
        //Create Word document.
        Document document = new Document();

        //Load the file from disk.
        document.loadFromFile("data/JAVATemplate_N.docx");

        //Create an instance of ToPdfParameterList.
        ToPdfParameterList toPdf = new ToPdfParameterList();

        //Set the user password for the resulted PDF file.
        toPdf.getPdfSecurity().encrypt("e-iceblue","test", PdfPermissionsFlags.Default, PdfEncryptionKeySize.Key_128_Bit);

        //Save to file.
        String result = "output/WordToPdfEncrypt.pdf";
        document.saveToFile(result, toPdf);
    }
}
