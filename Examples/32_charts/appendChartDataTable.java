import com.spire.doc.*;
import com.spire.doc.documents.Paragraph;
import com.spire.doc.fields.ShapeObject;
import com.spire.doc.fields.shapes.charts.Chart;
import com.spire.doc.fields.shapes.charts.ChartDataLabelCollection;
import com.spire.doc.fields.shapes.charts.ChartSeriesCollection;

public class appendChartDataTable {
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

                        // Call the method to add or update the chart data table
                        appendChartDataTable(chart);
                    }
                }
            }
        }

        document.saveToFile("appendChartDataTable.docx",FileFormat.Docx_2019);

        //Dispose the document
        document.dispose();
    }
    static void appendChartDataTable(Chart chart)
    {
        // Enable the display of the data table in the chart
        chart.getDataTable().setShow(true);

        // Show legend keys (symbols) in the data table
        chart.getDataTable().setShowLegendKeys(true);

        // Display horizontal borders between rows in the data table
        chart.getDataTable().setShowHorizontalBorder(true);

        // Display vertical borders between columns in the data table
        chart.getDataTable().setShowVerticalBorder(true);

        // Show an outline border around the entire data table
        chart.getDataTable().setShowOutlineBorder(true);
    }
}
