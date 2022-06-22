import com.spire.doc.*;
import com.spire.doc.documents.Paragraph;

import java.io.*;

public class getTextByStyleName {
    public static void main(String[] args) throws  Exception {

        String inputFile ="data/Template_N5.docx";
        String outputFile ="output/GetTextByStyleName.txt";

        //Load document from disk
        Document doc = new Document();
        doc.loadFromFile(inputFile);

        String text = "";

        //Loop through sections
        for (int i = 0; i < doc.getSections().getCount(); i ++)
        {
           Section section = doc.getSections().get(i);
           for(int j = 0 ; j < section.getParagraphs().getCount(); j ++)
           {
              Paragraph para = section.getParagraphs().get(j);
              String name = para.getStyleName();

               if (para.getStyleName().equals("Heading1"))
               {
                   //Write the text of paragraph
                   text += para.getText();
               }
           }
        }
        writeStringToTxt(text, outputFile);
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
