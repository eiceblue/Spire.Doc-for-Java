import com.spire.doc.*;
import com.spire.doc.documents.*;

public class removePageBreaks {
    public static void main(String[] args) {
        //Create Word document.
        Document document = new Document();

        //Load the file from disk.
        document.loadFromFile("data/Template_Docx_4.docx");

        //Traverse every paragraph of the first section of the document.
        for (int j = 0; j < document.getSections().get(0).getParagraphs().getCount(); j++) {
            Paragraph p = document.getSections().get(0).getParagraphs().get(j);

            //Traverse every child object of a paragraph.
            for (int i = 0; i < p.getChildObjects().getCount(); i++) {
                DocumentObject obj = p.getChildObjects().get(i);

                //Find the page break object.
                if (obj.getDocumentObjectType().equals(DocumentObjectType.Break)) {
                    Break b = (Break) obj;

                    //Remove the page break object from paragraph.
                    p.getChildObjects().remove(b);
                }
            }
        }

        String result = "output/result-removePageBreaks.docx";

        //Save to file.
        document.saveToFile(result, FileFormat.Docx_2013);
    }
}
