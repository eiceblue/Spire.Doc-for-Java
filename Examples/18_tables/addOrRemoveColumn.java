import com.spire.doc.*;

public class addOrRemoveColumn {
    public static void main(String[] args) {
    String input="data/tableSample.docx";
    String output="output/addOrRemoveColumn.docx";

    //load the document
    Document doc = new Document();
    doc.loadFromFile(input);

    //access the first section
    Section section = doc.getSections().get(0);

    //access the first table
    Table table = section.getTables().get(0);

    //add a blank column
    int columnIndex1 = 0;
    addColumn(table, columnIndex1);

    //remove a column
    int columnIndex2 = 2;
    removeColumn(table, columnIndex2);

    //save the document
    doc.saveToFile(output, FileFormat.Docx);
}
    private static void addColumn(Table table, int columnIndex)
    {
        for (int r = 0; r < table.getRows().getCount(); r++)
        {
            TableCell addCell = new TableCell(table.getDocument());
            table.getRows().get(r).getCells().insert(columnIndex, addCell);
        }
    }
    private static void removeColumn(Table table, int columnIndex)
    {
        for (int r = 0; r < table.getRows().getCount(); r++)
        {
            table.getRows().get(r).getCells().removeAt(columnIndex);
        }
    }
}
