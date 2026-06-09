import com.spire.doc.Document;
import com.spire.doc.FileFormat;
import com.spire.doc.documents.*;
import com.spire.doc.interfaces.IStyle;

import java.awt.*;

public class modifyCustomStyles {
    public static void main(String[] args) {
        // Create a new Document instance.
        Document document = new Document();

        // Load an existing Word document that contains custom styles from the specified relative file path.
        document.loadFromFile("Data/CustomStyles.docx");

        // Attempt to find a paragraph style named "MyParagraphStyle1" in the document's style collection.
        IStyle style1 = document.getStyles().findByName("MyParagraphStyle1");

        // Check if the style exists and is of type ParagraphStyle.
        if (style1 != null && style1 instanceof ParagraphStyle){
            // Cast the found style to ParagraphStyle for modification.
            ParagraphStyle cStyle = (ParagraphStyle)style1;

            // Set the font name of the paragraph style to "Arial".
            cStyle.getCharacterFormat().setFontName("Arial");

            // Set the font size to 18 points.
            cStyle.getCharacterFormat().setFontSize(18);

            // Disable bold formatting.
            cStyle.getCharacterFormat().setBold(false);

            // Change the text color to gray.
            cStyle.getCharacterFormat().setTextColor(Color.gray);
        }

        // Attempt to find a table style named "MyTableStyle1" in the document's style collection.
        IStyle style2 = document.getStyles().findByName("MyTableStyle1");

        // Check if the style exists and is of type TableStyle.
        if (style2 != null && style2 instanceof TableStyle) {
            // Cast the found style to TableStyle for modification.
            TableStyle tableStyle = (TableStyle)style2;

            // Set the border color of the table to gray.
            tableStyle.getBorders().setColor(Color.gray);

            // Change the border type to single line.
            tableStyle.getBorders().setBorderType(BorderStyle.Single);

            // Set the border line width to 1 point.
            tableStyle.getBorders().setLineWidth(1);

            // Apply row striping every 2 rows (i.e., stripe pattern repeats every two rows).
            tableStyle.setRowStripe(2);
        }

        // Save the modified document to a new file using the Word 2016 .docx format.
        document.saveToFile("ModifyCustomStyles.docx", FileFormat.Docx_2016);

        // Close the document to release associated resources.
        document.close();

        // Explicitly dispose of the document object to free memory immediately.
        document.dispose();
    }
}
