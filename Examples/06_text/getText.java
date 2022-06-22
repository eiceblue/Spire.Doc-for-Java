import com.spire.doc.*;
import java.io.*;

public class getText {
    public static void main(String[] args) throws Exception{
        //Load the document from disk.
        Document document = new Document();
        document.loadFromFile("data/ExtractText.docx");

        //Create a new TXT File to save extracted text
        String result = "output/extractText.txt";
        File file=new File(result);
        if(!file.exists()){
            file.delete();
        }
        file.createNewFile();
        FileWriter fw=new FileWriter(file,true);
        BufferedWriter bw=new BufferedWriter(fw);

        //Get text from document
        String text = document.getText();

        //Save extracted text
        bw.write(text);

        bw.flush();
        bw.close();
        fw.close();
    }
}
