import com.spire.doc.*;

public class keepSameFormat {
    public static void main(String[] args) {
        //Load the source document from disk
        Document srcDoc = new Document();
        srcDoc.loadFromFile("data/Template_N2.docx");

        //Load the destination document from disk
        Document destDoc = new Document();
        destDoc.loadFromFile("data/Template_N3.docx");

        //Keep same format of source document
        srcDoc.setKeepSameFormat(true);

        //Copy the sections of source document to destination document
        for (Object sectionObj : srcDoc.getSections()) {
            Section section=(Section)sectionObj;
            destDoc.getSections().add(section.deepClone());
        }

        //Save the Word document
        String output="output/keepSameFormat.docx";
        destDoc.saveToFile(output, FileFormat.Docx_2013);
    }
}
