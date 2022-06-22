import com.spire.doc.*;
import com.spire.doc.documents.Paragraph;
import com.spire.doc.fields.TextRange;

public class embedPrivateFont {
    public static void main(String[] args) {

        String inputFile="data/BlankTemplate.docx";
        String fontFile="data/PT_Serif-Caption-Web-Regular.ttf";
        String outputFile="output/embedPrivateFont.docx";

        Document doc = new Document();
        doc.loadFromFile(inputFile);

        //Get the first section and add a paragraph
        Section section = doc.getSections().get(0);
        Paragraph p = section.addParagraph();

        //Append text to the paragraph, then set the font name and font size
        TextRange range = p.appendText("Spire.Doc for Java is a professional Word Java library specifically designed for developers to create, read, write, convert and print Word document files from Java platform with fast and high quality performance.");
        range.getCharacterFormat().setFontName( "PT Serif Caption");
        range.getCharacterFormat().setFontSize(20);

        //Allow embedding font in document
        doc.setEmbedFontsInFile(true);

        //Embed private font from font file into the document
        doc.getPrivateFontList().add(new PrivateFontPath("PT Serif Caption",fontFile));

        //Save document
        doc.saveToFile(outputFile, FileFormat.Docx);
    }
}
