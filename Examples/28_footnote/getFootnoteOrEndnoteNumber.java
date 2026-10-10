import com.spire.doc.*;
import com.spire.doc.documents.*;
import com.spire.doc.fields.*;
import java.io.*;

public class getFootnoteOrEndnoteNumber {
    public static void main(String[] args) throws Exception {
        // Create a new Document object to load and process the Word file.
        Document doc = new Document();

        // Load the Word document containing footnotes and endnotes from the specified file path.
        doc.loadFromFile("Data/footnoteAndEndnote.docx");

        // Initialize a StringBuilder to store the extracted footnote and endnote information.
        StringBuilder stringBuilder = new StringBuilder();

        // Iterate through all sections in the document.
        for (int n = 0; n < doc.getSections().getCount(); n++) {

            // Get the current section by its index.
            Section s = doc.getSections().get(n);

            // Iterate through all paragraphs within the current section.
            for (int i = 0; i < s.getParagraphs().getCount(); i++) {

                // Get the current paragraph by its index.
                Paragraph para = s.getParagraphs().get(i);

                // Iterate through all child objects within the current paragraph.
                for (int j = 0, cnt = para.getChildObjects().getCount(); j < cnt; j++) {

                    // Get the current child object and cast it to a ParagraphBase type.
                    ParagraphBase pBase = (ParagraphBase) para.getChildObjects().get(j);

                    // Check if the current child object is a Footnote (which also includes Endnotes).
                    if (pBase instanceof Footnote) {

                        // Cast the child object to a Footnote instance.
                        Footnote fn = (Footnote) pBase;

                        // Check if the current note is specifically a Footnote.
                        if (fn.getFootnoteType() == FootnoteType.Footnote) {

                            // Create a StringBuilder to accumulate the text content of the footnote.
                            StringBuilder fnText = new StringBuilder();

                            // Iterate through all paragraphs in the footnote's text body.
                            for (int k = 0; k < fn.getTextBody().getParagraphs().getCount(); k++) {

                                // Append the text of each paragraph to the footnote text builder.
                                fnText.append(fn.getTextBody().getParagraphs().get(k).getText());
                            }

                            // Append the extracted footnote content and its ID to the main result string.
                            stringBuilder.append("Footnote content:" + fnText.toString() + "\nFootnote ID：" + fn.getId() + "\n");
                        }

                        // Check if the current note is specifically an Endnote.
                        if (fn.getFootnoteType() == FootnoteType.Endnote) {

                            // Create a StringBuilder to accumulate the text content of the endnote.
                            StringBuilder enText = new StringBuilder();

                            // Iterate through all paragraphs in the endnote's text body.
                            for (int k = 0; k < fn.getTextBody().getParagraphs().getCount(); k++) {

                                // Append the text of each paragraph to the endnote text builder.
                                enText.append(fn.getTextBody().getParagraphs().get(k).getText());
                            }

                            // Append the extracted endnote content and its ID to the main result string.
                            stringBuilder.append("Endnote content:" + enText.toString() + "\nEndnote ID：" + fn.getId() + "\n");
                        }
                    }
                }
            }
        }

        // Write the collected footnote and endnote information to a specified text file.
        writeStringToTxt(stringBuilder.toString(), "GetFootnoteOrEndnoteNumber.txt");

        doc.close();
        doc.dispose();
    }
    // Custom method to write a string to a text file
    public static void writeStringToTxt(String content, String txtFileName) throws IOException {
        FileWriter fWriter = new FileWriter(txtFileName, true);
        try {
            // Write the content to the text file
            fWriter.write(content);
        } catch (IOException ex) {
            ex.printStackTrace();
        } finally {
            try {
                // Flush and close the FileWriter
                fWriter.flush();
                fWriter.close();
            } catch (IOException ex) {
                ex.printStackTrace();
            }
        }
    }
}
