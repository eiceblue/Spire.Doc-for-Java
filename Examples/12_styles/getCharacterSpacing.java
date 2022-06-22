import com.spire.doc.*;
import com.spire.doc.documents.Paragraph;
import com.spire.doc.fields.TextRange;

public class getCharacterSpacing {
    public static void main(String[] args) {

        String inputFile ="data/Insert.docx";

        //Create a document
        Document document = new Document();

        //Load the document from disk.
        document.loadFromFile(inputFile);

        Section section = document.getSections().get(0);
        Paragraph p = section.getParagraphs().get(0);

        //Traverse the child objects
        for (int j = 0; j < p.getChildObjects().getCount(); j ++)
        {
            if ( p.getChildObjects().get(j) instanceof TextRange )
            {
                TextRange tr = (TextRange) p.getChildObjects().get(j);

                String fontName = tr.getCharacterFormat().getFontName();
                float fontSpacing = tr.getCharacterFormat().getCharacterSpacing();
            }
        }

    }
}
