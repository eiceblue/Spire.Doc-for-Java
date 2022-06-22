import com.spire.doc.*;
import com.spire.doc.collections.TableCollection;

import java.io.*;

public class getRowCellIndex {
    public static void main(String[] args) throws IOException {
        //Load Word from disk
        Document doc = new Document();
        doc.loadFromFile("data/ReplaceTextInTable.docx");

        //Get the first section
        Section section = doc.getSections().get(0);

        //Get the first table in the section
        Table table = section.getTables().get(0);

        StringBuilder content = new StringBuilder();

        //Get table collections
        TableCollection collections = section.getTables();

        //Get the table index
        int tableIndex = collections.indexOf(table);

        //Get the index of the last table row
        TableRow row = table.getLastRow();
        int rowIndex = row.getRowIndex();

        //Get the index of the last table cell
        TableCell cell = (TableCell)row.getLastChild();
        int cellIndex = cell.getCellIndex();

        //Append these information into content
        content.append("Table index is " + tableIndex);
        content.append("\r\n");
        content.append("Row index is " + rowIndex);
        content.append("\r\n");
        content.append("Cell index is " + cellIndex);

        //Save to txt file
        String output = "output/GetRowCellIndex_out.txt";
        writeStringToTxt(content.toString(),output);
    }
    public static void writeStringToTxt(String content, String txtFileName) throws IOException {
        File file=new File(txtFileName);
        if (file.exists())
        {
            file.delete();
        }
        FileWriter fWriter = new FileWriter(txtFileName, true);
        try {
            fWriter.write(content);
        } catch (IOException ex) {
            ex.printStackTrace();
        } finally {
            try {
                fWriter.flush();
                fWriter.close();
            } catch (IOException ex) {
                ex.printStackTrace();
            }
        }
    }
}
