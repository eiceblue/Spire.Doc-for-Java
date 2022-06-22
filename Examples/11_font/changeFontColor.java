import com.spire.doc.*;
import com.spire.doc.documents.*;
import com.spire.doc.fields.TextRange;

import java.awt.*;

public class changeFontColor {
    public static void main(String[] args) {
        String inputFile="data/Sample.docx";
        String outputFile="output/ChangeFontColor.docx";

        Document doc = new Document();
        doc.loadFromFile(inputFile);

        //Get the first section and first paragraph
        Section section = doc.getSections().get(0);
        Paragraph p1 = section.getParagraphs().get(0);

        //Iterate through the childObjects of the paragraph 1
        for (int i = 0; i < p1.getChildObjects().getCount(); i ++)
        {
            if ( p1.getChildObjects().get(i) instanceof TextRange)
            {
                TextRange tr = (TextRange) p1.getChildObjects().get(i);
                tr.getCharacterFormat().setTextColor(Color.red);
            }
        }

        //Get the second paragraph
        Paragraph p2 = section.getParagraphs().get(1);

        //Iterate through the childObjects of the paragraph 2
        for (int j = 0; j < p2.getChildObjects().getCount(); j ++)
        {
            if ( p2.getChildObjects().get(j) instanceof TextRange)
            {
                TextRange tr = (TextRange) p2.getChildObjects().get(j);
                tr.getCharacterFormat().setTextColor(Color.GRAY);
            }
        }
        doc.saveToFile(outputFile, FileFormat.Docx);
    }
}
