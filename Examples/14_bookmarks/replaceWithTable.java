import com.spire.data.table.*;
import com.spire.doc.*;
import com.spire.doc.documents.*;

public class replaceWithTable {
    public static void main(String[] args) throws Exception {
        String input = "data/bookmarks.docx";
        String output = "output/replaceWithTable.docx";

        //Load the document
        Document doc = new Document();
        doc.loadFromFile(input);

        //Create a table
        Table table = new Table(doc, true);

        //Create datatable
        DataTable dataTable = new DataTable();
        Class stringClass = String.class;
        dataTable.getColumns().add("id",stringClass);
        dataTable.getColumns().add("name",stringClass);
        dataTable.getColumns().add("job",stringClass);
        dataTable.getColumns().add("email",stringClass);
        dataTable.getColumns().add("salary",stringClass);

        //Table data
        String[][] data =
                {
                        new String[]{"Name", "Capital", "Continent", "Area", "Population"},
                        new String[]{"Argentina", "Buenos Aires", "South America", "2777815", "32300003"},
                        new String[]{"Bolivia", "La Paz", "South America", "1098575", "7300000" },
                        new String[]{"Brazil", "Brasilia", "South America", "8511196", "150400000"},
                };
        table.resetCells(data.length,dataTable.getColumns().size());

        //Fill the data to Table
        for (int i = 0; i < data.length; i++)
        {
            for (int j = 0; j < dataTable.getColumns().size(); j++)
            {
                table.getRows().get(i).getCells().get(j).addParagraph().appendText(data[i][j]);
            }
        }
        //Get the specific bookmark by its name
        BookmarksNavigator navigator = new BookmarksNavigator(doc);
        navigator.moveToBookmark("Test2");

        //Create a TextBodyPart instance and add the table to it
        TextBodyPart part = new TextBodyPart(doc);
        part.getBodyItems().add(table);

        //Replace the current bookmark content with the TextBodyPart object
        navigator.replaceBookmarkContent(part);

        //Save the document
        doc.saveToFile(output, FileFormat.Docx);
    }
}
