import com.spire.doc.*;
import com.spire.doc.documents.*;
import com.spire.doc.fields.*;
import com.spire.doc.interfaces.*;

import java.awt.*;

public class addCustomStyles {
    public static void main(String[] args) {
        // Create a new instance of a Word document.
        Document doc = new Document();

        // Add a new section to the document.
        Section section = doc.addSection();

        // Add a new paragraph to the section.
        Paragraph paragraph = section.addParagraph();

        // Append initial text to the paragraph and obtain a reference to the resulting text range.
        TextRange textRange = paragraph.appendText("Spire.Doc for Java ");

        // Create and register a new character-style named "MyCharacterStyle1", then cast it to ICharacterStyle.
        ICharacterStyle characterStyle = (ICharacterStyle)doc.getStyles().add(StyleType.Character_Style, "MyCharacterStyle1");

        // Set the font name of the character style to Arial.
        characterStyle.getCharacterFormat().setFontName("Arial");

        // Set the font size of the character style to 14 points.
        characterStyle.getCharacterFormat().setFontSize(14);

        // Enable bold formatting for the character style.
        characterStyle.getCharacterFormat().setBold(true);

        // Set the text color of the character style to blue.
        characterStyle.getCharacterFormat().setTextColor(Color.blue);

        // Apply the custom character style to the previously created text range.
        textRange.applyStyle(characterStyle.getName());

        // Append additional descriptive text to the same paragraph.
        paragraph.appendText("is a professional Java Word API that enables Java applications to create, convert, manipulate and print Word documents without using Microsoft Office.");

        // Create and register a new paragraph-style named "MyParagraphStyle1", then cast it to ParagraphStyle.
        ParagraphStyle heading1Style = (ParagraphStyle)doc.getStyles().add(StyleType.Paragraph_Style, "MyParagraphStyle1");

        // Set the paragraph alignment to justified (text aligned evenly along both left and right margins).
        heading1Style.getParagraphFormat().setHorizontalAlignment(HorizontalAlignment.Justify);

        // Set the line spacing of the paragraph style to 18 points.
        heading1Style.getParagraphFormat().setLineSpacing(18);

        // Set the font name for the paragraph style.
        heading1Style.getCharacterFormat().setFontName("Calibri");

        // Set the font size for the paragraph style.
        heading1Style.getCharacterFormat().setFontSize(12);

        // Set a custom text color using RGB values.
        heading1Style.getCharacterFormat().setTextColor(new Color(42, 123, 136));

        // Add another new paragraph to the section.
        paragraph = section.addParagraph();
        paragraph.appendText("A plenty of Word document processing tasks can be performed by Spire.Doc for Java, such as create, read, edit, convert and print Word documents, insert image, add header and footer, create table, add form field and mail merge field, add bookmark and watermark, add hyperlink, set background color/image, add footnote and endnote, encrypt Word documents. ");

        // Apply the custom paragraph style to the previously created paragraph.
        paragraph.applyStyle(heading1Style);

        // Add a table to the section.
        Table table = section.addTable();
        // Initialize the table with 15 rows and 4 columns.
        table.resetCells(15, 4);
        // Loop through each row of the table.
        for (int i = 0; i < 15; i++)
        {
            // Get the current row.
            TableRow row = table.getRows().get(i);
            // Loop through each cell in the current row.
            for (int j = 0; j < 4; j++)
            {
                // Get the current cell.
                TableCell cell = row.getCells().get(j);
                // Add a paragraph to the cell and insert placeholder text.
                // The text alternates between "Start" and "Continuation" based on column index.
                String rowType = (j % 4 == 0) ? "Start" : "Continuation";
                cell.addParagraph().setText(String.format("Row %s", rowType));
            }
        }

        // Create a custom table style named "MyTableStyle1".
        TableStyle tableStyle = (TableStyle)doc.getStyles().add(StyleType.Table_Style, "MyTableStyle1");
        // Set the border color for the table.
        tableStyle.getBorders().setColor(Color.black);
        // Set the border type to double line.
        tableStyle.getBorders().setBorderType(BorderStyle.Double);
        // Apply row striping every 3 rows.
        tableStyle.setRowStripe(3);
        // Set background color for odd-numbered striped rows.
        tableStyle.getConditionalStyles().get(TableConditionalStyleType.OddRowStripe.getValue()).getShading().setBackgroundPatternColor(new Color(173, 216, 230));
        // Set background color for even-numbered striped rows.
        tableStyle.getConditionalStyles().get(TableConditionalStyleType.EvenRowStripe.getValue()).getShading().setBackgroundPatternColor(new Color(224, 255, 255));
        // Apply column striping every 1 column (i.e., every other column).
        tableStyle.setColumnStripe(1);
        // Set background color for even-numbered striped columns.
        tableStyle.getConditionalStyles().get(TableConditionalStyleType.EvenColumnStripe.getValue()).getShading().setBackgroundPatternColor(new Color(255, 255, 224));
        // Apply the custom table style to the table.
        table.applyStyle(tableStyle);

        // Enable column striping in the table's formatting options.
        table.getFormat().setStyleOptions(TableStyleOptions.ColumnStripe);

        // Save the document to a file in .docx format.
        doc.saveToFile("AddCustomStyles.docx", FileFormat.Docx);

        // Close the document to release resources.
        doc.close();
        // Explicitly dispose of the document object.
        doc.dispose();
    }
}
