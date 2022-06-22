import com.spire.doc.*;

public class copyContentToAnotherDoc {
    public static void main(String[] args) {
        //Initialize a new object of Document class and load the source document.
        Document sourceDoc = new Document();
        sourceDoc.loadFromFile("data/Template_Docx_1.docx");

        //Initialize another object to load target document.
        Document destinationDoc = new Document();
        destinationDoc.loadFromFile("data/Target.docx");

        //Copy content from source file and insert them to the target file.
        for (Object sectionObj : sourceDoc.getSections()) {
            Section sec=(Section)sectionObj;
            for (Object docObj : sec.getBody().getChildObjects()) {
                DocumentObject obj=(DocumentObject)docObj;
                destinationDoc.getSections().get(0).getBody().getChildObjects().add(obj.deepClone());
            }
        }

        String result = "output/copyContentToAnotherWord.docx";

        //Save to file.
        destinationDoc.saveToFile(result, FileFormat.Docx_2013);
    }
}
