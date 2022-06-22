import com.spire.doc.*;
import com.spire.doc.reporting.*;
import java.awt.*;

public class alternateRows {

    public static void main(String[] args) throws Exception {

        String data = "data/dataSample.xml";
        String input = "data/ExecuteMailMergeSample.doc";
        String output = "output/alternateRows.doc";

        Document doc = new Document();
        doc.loadFromFile(input);

        doc.getMailMerge().MergeField = new MergeFieldEventHandler() {

            public void invoke(Object sender, MergeFieldEventArgs args) {
                mailMerge_MergeField(sender, args);
            }
        };

        //Mail merge
        doc.getMailMerge().executeWidthRegion(data);

        doc.saveToFile(output, FileFormat.Doc);

    }
    static int rowIndex = 0;
    private static void mailMerge_MergeField(Object sender, MergeFieldEventArgs args) {
        {
            // Catch the beginning of a new row.
            if (args.getCurrentMergeField().getFieldName().equals("Name")) {
                // Set the color depending on whether the row number is even or odd.
                Color rowColor;
                if (rowIndex % 2 == 0)
                    rowColor = Color.gray;
                else
                    rowColor = Color.lightGray;

                TableCell cell = ( TableCell)args.getCurrentMergeField().getOwnerParagraph().getOwner();
                TableRow row = cell.getOwnerRow();

                row.getRowFormat().setBackColor(rowColor);
                rowIndex++;
            }
        }
    }
}
