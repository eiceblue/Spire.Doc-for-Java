import com.spire.doc.*;
import com.spire.doc.documents.*;
import com.spire.doc.fields.*;

import java.io.*;

public class extractTextFromTextBoxes {
    public static void main(String[] args) throws IOException {

        //Create a Word document.
        Document document = new Document();

        //Load the file from disk.
        document.loadFromFile("data/extractTextFromTextBoxes.docx");

        //Set output path
        String output = "output/extractTextFromTextBoxes.txt";

        //Verify whether the document contains a textbox or not.
        if (document.getTextBoxes().getCount() > 0) {
            File file = new File(output);
            FileWriter fw = new FileWriter(file, true);
            BufferedWriter bw = new BufferedWriter(fw);
            //Traverse the document.
            for (int s = 0; s < document.getSections().getCount(); s++) {
                Section section = document.getSections().get(s);
                for (int i = 0; i < section.getParagraphs().getCount(); i++) {
                    Paragraph p = section.getParagraphs().get(i);
                    for (int j = 0; j < p.getChildObjects().getCount(); j++) {
                        DocumentObject obj = p.getChildObjects().get(j);
                        if (obj.getDocumentObjectType() == DocumentObjectType.Text_Box) {
                            TextBox textbox = (TextBox) obj;
                            for (int k = 0; k < textbox.getChildObjects().getCount(); k++) {
                                DocumentObject objt = textbox.getChildObjects().get(k);
                                //Extract text from paragraph in TextBox.
                                if (objt.getDocumentObjectType() == DocumentObjectType.Paragraph) {
                                    bw.write(((Paragraph) objt).getText());
                                }

                                //Extract text from Table in TextBox.
                                if (objt.getDocumentObjectType() == DocumentObjectType.Table) {
                                    Table table = (Table) objt;
                                    ExtractTextFromTables(table, bw);
                                }
                            }
                        }
                    }
                }
            }
            bw.flush();
            bw.close();
            fw.close();
        }
    }

    //Extract text from Table .
    static void ExtractTextFromTables(Table table, BufferedWriter sw) throws IOException {
        for (int i = 0; i < table.getRows().getCount(); i++) {
            TableRow row = table.getRows().get(i);
            for (int j = 0; j < row.getCells().getCount(); j++) {
                TableCell cell = row.getCells().get(i);
                for (int k = 0; k < cell.getParagraphs().getCount(); k++) {
                    Paragraph paragraph = cell.getParagraphs().get(k);
                    sw.write(paragraph.getText());
                }
            }
        }
    }
}
