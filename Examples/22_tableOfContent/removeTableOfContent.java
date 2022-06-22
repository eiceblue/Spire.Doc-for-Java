import com.spire.doc.*;

import java.util.regex.*;

public class removeTableOfContent {
    public static void main(String[] args) {
        //Create a document
        Document document = new Document();

        //Load the document from disk.
        document.loadFromFile("data/tableOfContent.docx");

        //Get the first body from the first section
        Body body = document.getSections().get(0).getBody();

        //Remove TOC from first body
        String pattern = "TOC\\w+";
        for (int i = 0; i < body.getParagraphs().getCount(); i++) {
            String styleName = body.getParagraphs().get(i).getStyleName();
            if (Pattern.matches(pattern, styleName)) {
                body.getParagraphs().removeAt(i);
                i--;
            }
        }

        //Save the document.
        String output = "output/removeTableOfContent.docx";
        document.saveToFile(output, FileFormat.Docx_2013);
    }
}
