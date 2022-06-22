import com.spire.doc.*;
import java.io.*;

public class getBookmarks {
    public static void main(String[] args) throws IOException {
        String input = "data/bookmarks.docx";
        String output = "output/getBookmarks.txt";

        //Create word document
        Document document = new Document();

        //Load the document from disk.
        document.loadFromFile(input);

        //Get the bookmark by index.
        Bookmark bookmark1 = document.getBookmarks().get(0);

        //Get the bookmark by name.
        Bookmark bookmark2 = document.getBookmarks().get("Test2");

        //Set string format for displaying
        String result = String.format("The bookmark obtained by index is " + bookmark1.getName() + ".\r\nThe bookmark obtained by name is " + bookmark2.getName() + ".\n");

        //create a new TXT file to save the text
        writeStringToTxt(result,output);
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
