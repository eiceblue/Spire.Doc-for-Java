import com.spire.doc.*;

public class preserveTheme {
    public static void main(String[] args) {
        //Load the source document
        String input = "data/Theme.docx";
        Document doc = new Document();
        doc.loadFromFile(input);

        //Create a new Word document
        Document newWord = new Document();

        //Clone default style, theme, compatibility from the source file to the destination document
        doc.cloneDefaultStyleTo(newWord);
        doc.cloneThemesTo(newWord);
        doc.cloneCompatibilityTo(newWord);

        //Add the cloned section to destination document
        newWord.getSections().add(doc.getSections().get(0).deepClone());

        //Save and launch document
        String output = "output/preserveTheme.docx";
        newWord.saveToFile(output, FileFormat.Docx_2013);
    }
}
