import com.spire.doc.*;

public class merge {
    public static void main(String[] args) {
        //Create word document and the file form disk
        Document document = new Document();
        document.loadFromFile("data/Summary_of_Science.doc", FileFormat.Doc);

        //Create word document and the file form disk
        Document documentMerge = new Document();
        documentMerge.loadFromFile("data/Bookmark.docx", FileFormat.Docx);

        //Merge files
        for (Object sectionObj : documentMerge.getSections()) {
            Section section=(Section)sectionObj;
            document.getSections().add(section.deepClone());
        }

        //Save as docx file.
        document.saveToFile("output/merge.docx", FileFormat.Docx_2013);
    }
}
