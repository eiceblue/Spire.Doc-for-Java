import com.spire.doc.*;
import com.spire.doc.documents.*;
import com.spire.doc.fields.*;

public class removeFootnote {
    public static void main(String[] args) {
        Document document = new Document();
        document.loadFromFile("data/footnote.docx");
        Section section = document.getSections().get(0);

        //Traverse paragraphs in the section and find the footnote
        for (int i = 0; i < section.getParagraphs().getCount(); i++)
        {
            Paragraph para = section.getParagraphs().get(i);
            int index = -1;
            for (int j = 0, cnt = para.getChildObjects().getCount(); j < cnt; j++)
            {
                ParagraphBase pBase = (ParagraphBase)para.getChildObjects().get(j);
                if (pBase instanceof Footnote)
                {
                    index = j;
                    break;
                }
            }

            if (index > -1)
                //Remove the footnote
                para.getChildObjects().removeAt(index);
        }

        //Save the document
        String output = "output/removeFootnote.docx";
        document.saveToFile(output, FileFormat.Docx);

    }
}
