import com.spire.doc.*;
import com.spire.doc.documents.*;
import com.spire.doc.fields.*;
import com.spire.doc.fields.shapes.charts.*;

import java.awt.*;

public class appendChartAxis {
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
                    if (obj instanceof ShapeObject)
                    {
                        // Cast the object to a ShapeObject
                        ShapeObject shape = (ShapeObject)obj;

                        // Get the chart from the shape
                        Chart chart = shape.getChart();

                        // Call the method to add or update the chart axis
                        appendChartAxis(chart);
                    }
                }
            }
        }

        document.saveToFile("appendChartAxis.docx", FileFormat.Docx_2019);

        //Dispose the document
        document.dispose();
    }
    static void appendChartAxis(Chart chart)
    {
        for (int i = 0; i < chart.getAxes().getCount(); i++)
        {
            if (i == 0)
            {
                chart.getAxes().get(i).setCategoryType(AxisCategoryType.Category);
                chart.getAxes().get(i).getBounds().setMaximum(new AxisBound(5));
                chart.getAxes().get(i).getBounds().setMinimum(new AxisBound(0));
                chart.getAxes().get(i).getUnits().setMajor(1);
                chart.getAxes().get(i).getUnits().setMajorTimeUnit(AxisTimeUnit.Auto);
                chart.getAxes().get(i).getUnits().setMinor(1);
                chart.getAxes().get(i).getUnits().setMinorTimeUnit(AxisTimeUnit.Days);
                chart.getAxes().get(i).hasMajorGridlines(false);
                chart.getAxes().get(i).hasMinorGridlines( true);
                chart.getAxes().get(i).getLabels().isAutoSpacing(false);
                chart.getAxes().get(i).getLabels().setSpacing(1);
                chart.getAxes().get(i).getLabels().setOffset(1);
                chart.getAxes().get(i).getLabels().setPosition(AxisTickLabelPosition.Low);
                chart.getAxes().get(i).setReverseOrder(true);
                chart.getAxes().get(i).getTitle().setText("x-axis");
                chart.getAxes().get(i).getTitle().setShow(true);
                chart.getAxes().get(i).getTitle().setOverlay(true);
            }
            else if (i == 1)
            {
                chart.getAxes().get(i).setCategoryType(AxisCategoryType.Category);
                chart.getAxes().get(i).getUnits().isMajorAuto(true);
                chart.getAxes().get(i).getUnits().isMinorAuto( true);
                chart.getAxes().get(i).getBounds().setLogBase(10);
                chart.getAxes().get(i).hasMajorGridlines(true);
                chart.getAxes().get(i).hasMinorGridlines(  false);
                chart.getAxes().get(i).setReverseOrder(false);
                chart.getAxes().get(i).getLabels().isAutoSpacing(true);
                chart.getAxes().get(i).getTitle().setText("y-axis");
                chart.getAxes().get(i).getTitle().setShow( true);
                chart.getAxes().get(i).getTitle().setOverlay(true);
            }
            else
            {
                chart.getAxes().get(i).getTitle().setText( "z-axis");
                chart.getAxes().get(i).getTitle().setShow( true);
                chart.getAxes().get(i).getTitle().setOverlay(false);
            }
            chart.getAxes().get(i).getLabels().setAlignment(LabelAlignment.Left);
            chart.getAxes().get(i).getUnits().setBaseTimeUnit(AxisTimeUnit.Auto);
            chart.getAxes().get(i).setAxisBetweenCategories(true);
            chart.getAxes().get(i).getDisplayUnits().setCustomUnit(1);
            chart.getAxes().get(i).getDisplayUnits().setUnit( AxisBuiltInUnit.Custom);
            chart.getAxes().get(i).getDisplayUnits().setShowLabel(true);
            chart.getAxes().get(i).getTickMarks().setSpacing(1);
            chart.getAxes().get(i).getTickMarks().setMajor(AxisTickMark.None);
            chart.getAxes().get(i).getTickMarks().setMinor(AxisTickMark.Inside);
            chart.getAxes().get(i).getTitle().getCharacterFormat().setFontSize(8);
            chart.getAxes().get(i).getTitle().getCharacterFormat().setTextColor(Color.red);
            chart.getAxes().get(i).getTitle().getCharacterFormat().setBold(true);
        }
    }
}
