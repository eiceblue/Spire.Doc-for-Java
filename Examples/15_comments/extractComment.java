import com.spire.doc.*;
import com.spire.doc.documents.*;
import com.spire.doc.fields.*;
import java.io.*;

public class extractComment {
    public static void main(String[] args) throws IOException {
        String input = "data/commentSample.docx";
        String output = "output/extractComment.txt";

        //Load the document
        Document doc = new Document();
        doc.loadFromFile(input);

        //Traverse all comments
        for (int i=0;i< doc.getComments().getCount();i++)
        {
            Comment comment =doc.getComments().get(i);
            for (int j=0;j< comment.getBody().getParagraphs().getCount();j++)
            {
                Paragraph para = comment.getBody().getParagraphs().get(j);
                //Get comment
                String result =para.getText()+"\r\n";
                //create a new TXT file to save the text
                writeStringToTxt(result,output);
            }
        }
    }
    public static void writeStringToTxt(String content, String txtFileName) throws IOException {
        FileWriter fWriter= new FileWriter(txtFileName,true);
        try {
            fWriter.write(content);
        }catch(IOException ex){
            ex.printStackTrace();
        }finally{
            try{
                fWriter.flush();
                fWriter.close();
            } catch (IOException ex) {
                ex.printStackTrace();
            }
        }
    }
}
