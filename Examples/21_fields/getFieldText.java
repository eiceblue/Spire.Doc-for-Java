import com.spire.doc.Document;
import com.spire.doc.collections.FieldCollection;
import com.spire.doc.fields.Field;
import java.io.*;

public class getFieldText {
    public static void main(String[] args) throws IOException {
        StringBuilder sb = new StringBuilder();

        //Open a Word document
        Document document = new Document("data/GetFieldText.docx");

        //Get all fields in document
        FieldCollection fields = document.getFields();

        for (Field field : (Iterable<Field>)fields)
        {
            //Get field text
            String fieldText = field.getFieldText();
            sb.append("The field text is \""+fieldText + "\".\r\n");
        }

        //Write to txt file
        String output="output/getFieldText.txt";
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
