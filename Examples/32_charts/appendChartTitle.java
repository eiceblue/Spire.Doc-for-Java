import com.spire.doc.*;
import com.spire.doc.documents.*;
import com.spire.doc.fields.ShapeObject;
import com.spire.doc.fields.shapes.charts.Chart;
import com.spire.doc.fields.shapes.charts.ChartTitle;

import java.awt.*;

public class appendChartTitle {
    public static void main(String[] args) {

         //Create word document
        Document document = new Document();

        //Load the file from disk.
        document.loadFromFile("data\\ChartTemplate.docx");

        // Loop through all sections in the document
        for (int i = 0; i < document.getSections().getCount(); i++) {
            // Loop through all paragraphs in the current section
            for (int j = 0; j < document.getSections().get(i).getParagraphs().getCount(); j++) {
                // Get the current paragraph
                Paragraph paragraph = document.getSections().get(i).getParagraphs().get(j);

                // Loop through all child objects in the paragraph
                for (int k = 0; k < paragraph.getChildObjects().getCount(); k++) {
                    DocumentObject obj = paragraph.getChildObjects().get(k);
                    // Check if the object is a shape (e.g., chart, etc.)
                    if (obj instanceof ShapeObject) {
                        // Cast the object to a ShapeObject
                        ShapeObject shape = (ShapeObject) obj;

                        // Get the chart from the shape
                        Chart chart = shape.getChart();

                        appendChartTitle(chart);
                    }
                }
            }
        }

        document.saveToFile("appendChartTitle.docx", FileFormat.Docx_2019);

        //Dispose the document
        document.dispose();
    }
    static void appendChartTitle(Chart chart)
    {
        // Get the chart's title object
        ChartTitle title = chart.getTitle();

        // Enable the display of the title
        title.setShow(true);

        // Disable overlay so the title does not overlap with the chart area
        title.setOverlay(false);

        // Set the text of the title
        title.setText("My Chart");

        // Set font size of the title
        title.getCharacterFormat().setFontSize(12);

        // Set the title text to bold
        title.getCharacterFormat().setBold(true);

        // Set the text color to blue
        title.getCharacterFormat().setTextColor(Color.blue);

        // Enable right-to-left text formatting (if needed for language)
        title.getCharacterFormat().setBidi(true);

        // Apply italic style to the title text
        title.getCharacterFormat().setItalic(true);

        // Set character spacing (tracking or kerning)
        title.getCharacterFormat().setCharacterSpacing(2);

        // Set underline color to red
        title.getCharacterFormat().setUnderlineColor(Color.red);

        // Set underline style to double line
        title.getCharacterFormat().setUnderlineStyle(UnderlineStyle.Double);

        // Set font name
        title.getCharacterFormat().setFontName("arial");

        // Enable all caps formatting
        title.getCharacterFormat().setAllCaps(true);

        // Enable shadow effect on the text
        title.getCharacterFormat().isShadow(true);

        // Set the position of the text baseline relative to normal
        title.getCharacterFormat().setPosition(3);
    }
}
