import com.spire.doc.Document;

import java.io.*;

public class getMergeFieldName {
    public static void main(String[] args) throws IOException {
        StringBuilder sb = new StringBuilder();

        //Open a Word document
        Document document = new Document("data/MailMerge.doc");

        //Get merge field name
        String[] fieldNames = document.getMailMerge().getMergeFieldNames();

        sb.append("The document has " + fieldNames.length + " merge fields.");
        sb.append(" The below is the name of the merge field:"+"\r\n");
        for(String name : fieldNames)
        {
            sb.append(name+"\r\n");
        }

        //Write the information to txt file
        String output="output/getMergeFieldName.txt";
        writeStringToTxt(sb.toString(),output);
    }
    public static void writeStringToTxt(String content, String txtFileName) throws IOException {
        File file=new File(txtFileName);
        if (file.exists())
        {
            file.delete();
        }
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
