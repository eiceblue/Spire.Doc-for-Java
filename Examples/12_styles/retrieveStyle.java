import com.spire.doc.*;
import com.spire.doc.documents.Paragraph;

import java.io.*;

public class retrieveStyle {
    public static void main(String[] args) throws  Exception {

        String inputFile ="data/Styles.docx";
        String outputFile ="output/retrieveStyle.txt";

        //Load document from disk
        Document doc = new Document();
        doc.loadFromFile(inputFile);

        String StyleName = "";

        //Loop through sections
        for (int i = 0; i < doc.getSections().getCount(); i ++)
        {
           Section section = doc.getSections().get(i);
           for(int j = 0 ; j < section.getParagraphs().getCount(); j ++)
           {
              Paragraph para = section.getParagraphs().get(j);
               StyleName  += para.getStyleName();
           }
        }
        writeStringToTxt(StyleName, outputFile);
    }

    public static void writeStringToTxt(String content, String txtFileName) throws IOException {
        FileWriter fWriter = new FileWriter(txtFileName, true);
        try {
            fWriter.write(content);
        } catch (IOException ex) {
            ex.printStackTrace();
        } finally {
            try {
                fWriter.flush();
                fWriter.close();
            } catch (IOException ex) {
                ex.printStackTrace();
            }
        }
    }
}
