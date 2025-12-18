import com.spire.doc.*;
import com.spire.doc.documents.*;
import com.spire.doc.fields.*;
import com.spire.doc.fields.shapes.charts.*;
import java.awt.*;

public class appendChartLegend {
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

                        appendChartLegend(chart);
                    }
                }
            }
        }

        document.saveToFile("appendChartLegend.docx", FileFormat.Docx_2019);

        //Dispose the document
        document.dispose();
    }
    static void appendChartLegend(Chart chart)
    {
        // Enable the legend display on the chart
        chart.getLegend().setShow(true);

        // Set the position of the legend to the left side of the chart
        chart.getLegend().setPosition(LegendPosition.Left);

        // Disable overlay mode so the legend does not overlap with the chart plot area
        chart.getLegend().setOverlay(false);

        // Set the font size of the legend text to 9 points
        chart.getLegend().getCharacterFormat().setFontSize(9);

        // Set the text color of the legend labels to blue
        chart.getLegend().getCharacterFormat().setTextColor(Color.blue);

        // Apply italic style to the legend text
        chart.getLegend().getCharacterFormat().setItalic(true);
    }
}
