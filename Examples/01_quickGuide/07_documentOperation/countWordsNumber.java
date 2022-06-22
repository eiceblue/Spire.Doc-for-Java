import com.spire.doc.*;
public class countWordsNumber {
    public static void main(String[] args) {
        Document document = new Document();
        //Load the file from disk.
        document.loadFromFile("data/Template_Docx_1.docx");

        //Count the number of words.
        System.out.println("CharCount: " + document.getBuiltinDocumentProperties().getCharCount());
        System.out.println("CharCountWithSpace: " + document.getBuiltinDocumentProperties().getCharCountWithSpace());
        System.out.println("WordCount: " + document.getBuiltinDocumentProperties().getWordCount());
    }
}
