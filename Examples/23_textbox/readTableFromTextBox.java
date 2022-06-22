import com.spire.doc.*;
import com.spire.doc.documents.*;
import com.spire.doc.fields.*;

import java.io.*;

public class readTableFromTextBox {
    public static void main(String[] args) throws IOException {
        //Load the document
        Document doc = new Document();
        doc.loadFromFile("data/textBoxTable.docx");

        //Get the first textbox
        TextBox textbox = doc.getTextBoxes().get(0);

        //Get the first table in the textbox
        Table table = textbox.getBody().getTables().get(0);

        //Set output path
        String output = "output/readTableFromTextBox.txt";
        File file = new File(output);
        if (!file.exists()) {
            file.delete();
        }
        file.createNewFile();
        FileWriter fw = new FileWriter(file, true);
        BufferedWriter bw = new BufferedWriter(fw);

        //Loop through the paragraphs of the table cells and extract them to a .txt file
        for (int i = 0; i < table.getRows().getCount(); i++) {
            TableRow row = table.getRows().get(i);
            for (int j = 0; j < row.getCells().getCount(); j++) {
                TableCell cell = row.getCells().get(j);
                for (int k = 0; k < cell.getParagraphs().getCount(); k++) {
                    Paragraph paragraph = cell.getParagraphs().get(k);
                    bw.write(paragraph.getText() + "\t");
                }
            }
            bw.write("\r\n");
        }

        bw.flush();
        bw.close();
        fw.close();
    }
}
