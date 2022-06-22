import com.spire.doc.*;
import com.spire.doc.collections.FormFieldCollection;
import java.io.*;

public class getFormFieldsCollection {
    public static void main(String[] args) throws IOException {
        StringBuilder sb = new StringBuilder();

        //Open a Word document
        Document document = new Document("data/FillFormField.doc");

        //Get the first section
        Section section = document.getSections().get(0);

        FormFieldCollection formFields = section.getBody().getFormFields();

        sb.append("The first section has " + formFields.getCount() + " form fields.");

        //Write the information to txt file
        String output="output/getFormFieldsCollection.txt";
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
