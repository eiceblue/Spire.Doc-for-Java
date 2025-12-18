import com.spire.doc.*;
import com.spire.doc.documents.*;
import java.awt.*;

public class setTableStyleAndBorder {
    public static void main(String[] args) {
        String inputFile="data/tableSample.docx";
        String outputFile="output/setTableStyleAndBorder.docx";

		// Create a new document object
		Document document = new Document();

		// Load the document from the specified input file
		document.loadFromFile(inputFile);

		// Get the first section of the document
		Section section = document.getSections().get(0);

		// Get the first table in the section
		Table table = section.getTables().get(0);

		// Apply the "Colorful_List" default table style to the table
		table.applyStyle(DefaultTableStyle.Colorful_List);

		// Set the right border of the table to a red hairline with a line width of 1.0F
		table.getFormat().getBorders().getRight().setBorderType(BorderStyle.Hairline);
		table.getFormat().getBorders().getRight().setLineWidth(1.0F);
		table.getFormat().getBorders().getRight().setColor(Color.RED);

		// Set the top border of the table to a green hairline with a line width of 1.0F
		table.getFormat().getBorders().getTop().setBorderType(BorderStyle.Hairline);
		table.getFormat().getBorders().getTop().setLineWidth(1.0F);
		table.getFormat().getBorders().getTop().setColor(Color.GREEN);

		// Set the left border of the table to a yellow hairline with a line width of 1.0F
		table.getFormat().getBorders().getLeft().setBorderType(BorderStyle.Hairline);
		table.getFormat().getBorders().getLeft().setLineWidth(1.0F);
		table.getFormat().getBorders().getLeft().setColor(Color.YELLOW);

		// Set the bottom border of the table to a dot-dash style
		table.getFormat().getBorders().getBottom().setBorderType(BorderStyle.Dot_Dash);

		// Set the vertical borders of the table to a dot style and horizontal borders to none
		table.getFormat().getBorders().getVertical().setBorderType(BorderStyle.Dot);
		table.getFormat().getBorders().getHorizontal().setBorderType(BorderStyle.None);
		table.getFormat().getBorders().getVertical().setColor(Color.ORANGE);

		// Save the modified document to the specified output file in DOCX format
		document.saveToFile(outputFile, FileFormat.Docx);

		// Dispose the document resources
		document.dispose();
    }
}
