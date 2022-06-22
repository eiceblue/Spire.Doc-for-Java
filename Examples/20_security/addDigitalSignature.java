import com.spire.doc.*;

public class addDigitalSignature {

    public static void main(String[] args) {
        //Input and output file paths
        String input="data/AddDigitalSignature.docx";
        String output="output/AddDigitalSignature_output.docx";
        //The path of certificate file
        String certificatePath="data/gary.pfx";
        //Secure password
        String securePwd="e-iceblue";
        //Create a new Document object
        Document doc=new Document();
        //Load docx from input path
        doc.loadFromFile(input);
        //Add digital signature and save the document
        doc.saveToFile(output, FileFormat.Docx,certificatePath, securePwd);
        //Close the document
        doc.close();
    }
}
