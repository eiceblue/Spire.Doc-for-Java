import com.spire.doc.*;

public class removeCustomPropertyFields {
    public static void main(String[] args) {
        //Create Word document.
        Document document = new Document();

        //Load the file from disk.
        document.loadFromFile("data/removeCustomPropertyFields.docx");

        //Get custom document properties object.
        CustomDocumentProperties cdp = document.getCustomDocumentProperties();
        int i1 = cdp.getCount();
        //Remove all custom property fields in the document.
        for (int i = 0; i < cdp.getCount(); ) {
            cdp.remove(cdp.get(i).getName());
        }

        document.isUpdateFields(true);

        //Save to file.
        String output = "output/removeCustomPropertyFields.docx";
        document.saveToFile(output, FileFormat.Docx_2013);
    }
}
