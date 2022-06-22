import com.spire.doc.*;
import com.spire.doc.documents.*;
import com.spire.doc.fields.*;
import java.io.*;

public class extractBookmarkText {
    public static void main(String[] args) throws IOException {
        String input = "data/bookmarks.docx";
        String output = "output/extractBookmarkText.txt";

        //Load Document
        Document doc = new Document();
        doc.loadFromFile(input);

        //Creates a BookmarkNavigator instance to access the bookmark
        BookmarksNavigator navigator = new BookmarksNavigator(doc);

        //Locate a specific bookmark by bookmark name
        navigator.moveToBookmark("Test2");
        TextBodyPart textBodyPart = navigator.getBookmarkContent();

        //Iterate through the items in the bookmark content to get the text
        for (int i = 0; i < textBodyPart.getBodyItems().getCount(); i++)
        {
            if (textBodyPart.getBodyItems().get(i) instanceof Paragraph)
            {
                Paragraph itemPara = (Paragraph)textBodyPart.getBodyItems().get(i);
                for (int j = 0; j < itemPara.getChildObjects().getCount(); j++)
                {
                    if (itemPara.getChildObjects().get(j) instanceof TextRange)
                    {
                        TextRange textrange = (TextRange)(itemPara.getChildObjects().get(j));
                        String text = textrange.getText();
                        //create a new TXT file to save the text
                        writeStringToTxt(text,output);
                    }
                }
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
