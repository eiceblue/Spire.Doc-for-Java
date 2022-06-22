import com.spire.doc.*;
import com.spire.doc.fields.FormField;

import java.io.*;

public class getFormFieldByName {
    public static void main(String[] args) throws IOException {
        StringBuilder sb = new StringBuilder();

        //Open a Word document
        Document document = new Document("data/FillFormField.doc");

        //Get the first section
        Section section = document.getSections().get(0);

        //Get form field by name
        FormField formField = section.getBody().getFormFields().get("email");

        //Append the file name and type in string builder
        sb.append("The name of the form field is " + formField.getName()+"\r\n");
        sb.append("The type of the form field is " + formField.getFormFieldType());

        //Write the information to txt file
        String output="output/getFormFieldByName.txt";
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
