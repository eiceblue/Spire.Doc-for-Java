import com.spire.doc.*;

public class updateFields {
    public static void main(String[] args) {
        //Open a Word document
        Document document = new Document("data/ifFieldSample.docx");

        //Update fields
        document.isUpdateFields(true);

        //Save doc file
        String output = "output/updateFields.docx";
        document.saveToFile(output, FileFormat.Docx_2013);
    }
}
