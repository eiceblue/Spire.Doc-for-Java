import com.spire.doc.*;
import com.spire.doc.documents.*;
import com.spire.doc.fields.shapes.charts.*;

public class saveChartToTemplate {
    public static void main(String[] args) {
        // Create a new instance of the Word document
        Document doc = new Document();

        // Add a new section to the document
        Section section = doc.addSection();

        // Add a new paragraph to a newly created section
        Paragraph paragraph = section.addParagraph();

        // Append a column chart to the paragraph and retrieve the Chart object
        Chart chart = (paragraph.appendChart(ChartType.Column, 400, 300)).getChart();

        // Save the chart as a template file (.crtx)
        chart.saveAsTemplate("SaveChartToTemplate.crtx");

        // Close the document and release associated resources
        doc.close();

        // Dispose of the document object to free up memory
        doc.dispose();
    }
}
