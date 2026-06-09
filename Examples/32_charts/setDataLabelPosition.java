import com.spire.doc.*;
import com.spire.doc.documents.Paragraph;
import com.spire.doc.fields.ShapeObject;
import com.spire.doc.fields.shapes.charts.*;

public class setDataLabelPosition {
    public static void main(String[] args) {
        // Create a new Word document instance
        Document doc = new Document();

        // Add a new section to the document
        Section section = doc.addSection();

        // Add a paragraph with the text "Center" as a title/label
        section.addParagraph().appendText("Center");

        // Add a new paragraph to hold the first chart
        Paragraph newPara = section.addParagraph();

        // Append a Pie chart to the paragraph and set its size (width: 500, height: 300)
        ShapeObject shape = newPara.appendChart(ChartType.Pie, 500, 300);

        // Get the Chart object from the created shape
        Chart chart = shape.getChart();

        // Enable data labels for the first data series in the pie chart
        chart.getSeries().get(0).hasDataLabels(true);

        // Configure the data labels to display the category name
        chart.getSeries().get(0).getDataLabels().setShowCategoryName(true);

        // Configure the data labels to display the numeric value
        chart.getSeries().get(0).getDataLabels().setShowValue(true);

        // Set the position of the data labels to the center of the pie slices
        chart.getSeries().get(0).getDataLabels().setPosition(ChartDataLabelPosition.Center);

        // Add another paragraph with the text "Left" as a title/label
        section.addParagraph().appendText("Left");

        newPara = section.addParagraph();

        // Append a Bubble chart to the same paragraph and set its size (width: 500, height: 300)
        ShapeObject shape2 = newPara.appendChart(ChartType.Bubble, 500, 300);

        // Get the Chart object from the second shape
        Chart chart2 = shape2.getChart();

        // Enable data labels for the first data series in the bubble chart
        chart2.getSeries().get(0).hasDataLabels(true);

        // Configure the data labels to display the category name
        chart2.getSeries().get(0).getDataLabels().setShowCategoryName(true);

        // Configure the data labels to display the numeric value
        chart2.getSeries().get(0).getDataLabels().setShowValue(true);

        // Set the position of the data labels to the left side
        chart2.getSeries().get(0).getDataLabels().setPosition(ChartDataLabelPosition.Left);

        // Define the output file name for saving the document
        String outputFile = "SetDataLabelPosition.docx";

        // Save the document to the specified file in Docx format
        doc.saveToFile(outputFile, FileFormat.Docx);

        // Close the document and release associated resources
        doc.close();

        // Dispose of the document object to free up memory
        doc.dispose();
    }
}
