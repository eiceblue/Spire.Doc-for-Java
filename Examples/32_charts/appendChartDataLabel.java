import com.spire.doc.*;
import com.spire.doc.documents.*;
import com.spire.doc.fields.*;
import com.spire.doc.fields.shapes.charts.*;

import java.awt.*;

public class appendChartDataLabel {
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
                        ChartSeriesCollection series = chart.getSeries();
                        ChartDataLabelCollection dataLabels = series.get(0).getDataLabels();
                        series.get(0).hasDataLabels(true);
                        appendChartDataLabel(dataLabels);
                    }
                }
            }
        }

        document.saveToFile("appendChartDataLabel.docx",FileFormat.Docx_2019);

        //Dispose the document
        document.dispose();
    }
    static void appendChartDataLabel(ChartDataLabelCollection dataLabels)
    {
        // Display the value (e.g., percentage or numerical value) on the data labels
        dataLabels.setShowValue(true);

        // Display the category name (e.g., the label for each chart segment)
        dataLabels.setShowCategoryName(true);

        // Display the series name (useful when multiple series are present)
        dataLabels.setShowSeriesName( true);

        // Show leader lines connecting the data labels to the chart elements
        dataLabels.setShowLeaderLines( true);

        // Set the separator between different label components (e.g., value and category)
        dataLabels.setSeparator(";");

        // Set the number format for the displayed values (thousands separator and zero decimals)
        dataLabels.getNumberFormat().setFormatCode("#,##0");

        // Set the font size of the data labels
        dataLabels.getCharacterFormat().setFontSize(8);

        // Make the text in the data labels bold
        dataLabels.getCharacterFormat().setBold(true);

        // Set the text color of the data labels to blue
        dataLabels.getCharacterFormat().setTextColor( Color.blue);

        // Set the border color of the characters in the data labels to blue
        dataLabels.getCharacterFormat().getBorder().setColor(Color.blue);

        // Enable right-to-left (RTL) text direction for languages like Arabic or Hebrew
        dataLabels.getCharacterFormat().setBidi(true);

        // Apply italic formatting to the text
        dataLabels.getCharacterFormat().setItalic( true);

        // Set the underline color to red
        dataLabels.getCharacterFormat().setUnderlineColor(Color.red);

        // Set the underline style to double line
        dataLabels.getCharacterFormat().setUnderlineStyle(UnderlineStyle.Double);

        // Set the font family for the data labels
        dataLabels.getCharacterFormat().setFontName("Arial");

        // Display all text in uppercase letters
        dataLabels.getCharacterFormat().setAllCaps(true);

        // Apply a shadow effect to the text
        dataLabels.getCharacterFormat().isShadow( true);
    }
}
