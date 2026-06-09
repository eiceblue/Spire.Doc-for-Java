import com.spire.doc.*;
import com.spire.doc.documents.*;
import com.spire.doc.fields.shapes.charts.*;

public class createCombinedChart {
    public static void main(String[] args) {
        // Initialize a new Document object
        Document doc = new Document();

        // Add a new section to the document and create a paragraph within that section
        Paragraph paragraph = doc.addSection().addParagraph();

        // Append a column chart (450x300 pixels) to the paragraph and retrieve the Chart object
        Chart chart = paragraph.appendChart(ChartType.Column, 450, 300).getChart();

        // Change the chart type of the series named "Series 3" to a Line chart and enable secondary axis if applicable
        chart.changeSeriesType("Series 3", ChartSeriesType.Line, true);

        // Define the output file name for the combined chart document
        String outputFile = "CombinedChart.docx";

        // Save the document to the specified file in DOCX 2019 format
        doc.saveToFile(outputFile, FileFormat.Docx_2019);

        // Close the document to release resources
        doc.close();

        // Dispose of the document object to free up memory
        doc.dispose();
    }
}
