import com.spire.doc.*;

public class cloneTable {
    public static void main(String[] args) {
        //Load the document
        String input = "data/TableTemplate.docx";
        Document doc = new Document();
        doc.loadFromFile(input);

        //Get the first section
        Section se = doc.getSections().get(0);

        //Get the first table
        Table original_Table =se.getTables().get(0);

        //Copy the existing table to copied_Table via Table.clone()
        Table copied_Table = original_Table.deepClone();
        String[] st = new String[] { "Spire.Presentation for Java", "A professional " +
                "PowerPoint® compatible library that enables developers to create, read, " +
                "write, modify, convert and Print PowerPoint documents in Java Applications" };
        //Get the last row of table
        TableRow lastRow = copied_Table.getRows().get(copied_Table.getRows().getCount()-1);
        //Change last row data
        for (int i = 0; i < lastRow.getCells().getCount() - 1; i++)
        {
            lastRow.getCells().get(i).getParagraphs().get(0).setText(st[i]);
        }
        //Add copied_Table in section
        se.getTables().add(copied_Table);

        //Save and launch document
        String output = "output/CloneTable.docx";
        doc.saveToFile(output, FileFormat.Docx);
    }
}
