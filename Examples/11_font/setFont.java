import com.spire.doc.*;
import com.spire.doc.FileFormat;
import com.spire.doc.documents.Paragraph;
import com.spire.doc.fields.TextRange;
import com.spire.doc.formatting.CharacterFormat;

public class setFont {
    public static void main(String[] args) {

        String inputFile="data/Template_Docx_1.docx";
        String outputFile="output/setFont.docx";

        Document doc = new Document();
        doc.loadFromFile(inputFile);

        //Get the first section
        Section s = doc.getSections().get(0);

        //Get the paragraph
        Paragraph p = s.getParagraphs().get(2);

        //Create a characterFormat object
        CharacterFormat format = new CharacterFormat(doc);
        //Set font
        format.setFontName("Arial");
        format.setFontSize(16);

        //Loop through the childObjects of paragraph
        for (int j = 0; j < p.getChildObjects().getCount(); j ++)
        {
            if ( p.getChildObjects().get(j) instanceof TextRange)
            {
                TextRange tr = (TextRange) p.getChildObjects().get(j);
                //Apply character format
                tr.applyCharacterFormat(format);
            }
        }

        doc.saveToFile(outputFile, FileFormat.Docx);
    }
}
