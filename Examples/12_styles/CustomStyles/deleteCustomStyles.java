import com.spire.doc.*;
import com.spire.doc.documents.Paragraph;

public class deleteCustomStyles {
    public static void main(String[] args) {
        // Create a new Document instance.
        Document document = new Document();

        // Load an existing Word document that contains custom styles from the specified file path.
        document.loadFromFile("Data/CustomStyles.docx");

        Section section;
        // Iterate through each section in the document.
        for (int i=0;i< document.getSections().getCount();i++) {

            section = document.getSections().get(i);

            // Iterate through each paragraph in the body of the current section.
            for(int j=0; j< section.getBody().getParagraphs().getCount();j++) {

                Paragraph paragraph = section.getBody().getParagraphs().get(j);
                // Remove the style associated with the current paragraph from the document's style collection.
                document.getStyles().get(paragraph.getStyleName()).removeSelf();
            }

            Table table;
            TableRow tableRow;
            TableCell tableCell;
            Paragraph cellParagraph;
            // Iterate through each table in the body of the current section.
            for(int k =0; k< section.getBody().getTables().getCount();k++) {

                table = section.getBody().getTables().get(k);
                
                // Remove the style applied to the table itself.
                document.getStyles().get(table.getFormat().getStyleName()).removeSelf();

                // Iterate through each row in the current table.
                for (int j = 0; j < table.getRows().getCount() ; j++) {
                    tableRow = table.getRows().get(j);

                    // Iterate through each cell in the current table row.
                    for (int l = 0; l < tableRow.getCells().getCount(); l++) {

                        tableCell = tableRow.getCells().get(l);

                        // Iterate through each paragraph inside the current table cell.
                        for (int m = 0; m < tableCell.getParagraphs().getCount(); m++) {
                            cellParagraph = tableCell.getParagraphs().get(m);

                            // Remove the style associated with the paragraph inside the table cell.
                            document.getStyles().get(cellParagraph.getStyleName()).removeSelf();
                        }
                    }
                }
            }
        }

        // Save the modified document to a new file with custom styles removed.
        document.saveToFile("DeleteCustomStyles.docx", FileFormat.Docx);

        // Close the document to release internal resources.
        document.close();

        // Explicitly dispose of the document object to free memory.
        document.dispose();
    }
}
