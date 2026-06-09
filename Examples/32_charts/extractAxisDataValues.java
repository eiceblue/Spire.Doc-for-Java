import com.spire.doc.*;
import com.spire.doc.documents.Paragraph;
import com.spire.doc.fields.ShapeObject;
import com.spire.doc.fields.shapes.charts.Chart;
import com.spire.doc.fields.shapes.charts.ChartSeries;
import com.spire.doc.fields.shapes.charts.ChartValue;

import java.io.BufferedWriter;
import java.io.File;
import java.io.FileWriter;

public class extractAxisDataValues {
    public static void main(String[] args) throws Exception{
        // Create a new Document object for the first document
        Document doc = new Document();

        // Load the Word document from the specified relative file path
        doc.loadFromFile("Data/ExtractAxisDataValues.docx");

        // Initialize a StringBuilder to store the extracted axis data values
        StringBuilder stringBuilder = new StringBuilder();

        Section section;
        Paragraph paragraph;
        // Loop through each section in the loaded document
        for(int i=0;i<doc.getSections().getCount();i++) {
            section = doc.getSections().get(i);

            // Loop through each paragraph within the current section
            for(int j=0;j<section.getParagraphs().getCount();j++) {

                paragraph = section.getParagraphs().get(j);
                // Iterate over all child objects contained in the paragraph
                for (int k = 0; k < paragraph.getChildObjects().getCount(); k++) {
                    // Get the current document object at index i
                    DocumentObject obj = paragraph.getChildObjects().get(k);

                    // Check if the current object is a ShapeObject (which can contain charts)
                    if (obj instanceof ShapeObject){
                        // Cast the object to a ShapeObject
                        ShapeObject shape = (ShapeObject)obj;

                        // Retrieve the Chart object from the shape
                        Chart chart = shape.getChart();

                        // Add a header line indicating the start of X-axis data extraction
                        stringBuilder.append("Obtain X-axis data values:\r\n");

                        // Loop through all the X-axis values in the chart
                        for (int x = 0; x < chart.getXValues().getCount(); x++) {
                            // Get the specific X-axis value at the current index
                            ChartValue xVal = chart.getXValues().get(x);

                            // Append the string representation of the X-value to the StringBuilder, followed by a space
                            stringBuilder.append(xVal.getStringValue() + " ");
                        }

                        // Get the first data series from the chart (index 0)
                        ChartSeries series = chart.getSeries().get(0);

                        // Add a new line and a header for Y-axis data extraction
                        stringBuilder.append("\r\nObtain Y-axis data values:");

                        // Iterate through all the Y-values in the selected data series
                        for(int y=0;y<series.getYValues().getCount();y++) {
                            ChartValue yVal = series.getYValues().get(y);
                            // Append the numeric value of the Y-data point to the StringBuilder, followed by a space
                            stringBuilder.append(yVal.getValue() + " ");
                        }
                    }
                }
            }
        }

        // Define the output file name/path for saving the extracted data
        String result = "ExtractAxisDataValues_out.txt";

        //Create a new TXT File to save extracted text
        File file=new File(result);

        //Determine if the file exists
        if(!file.exists()){
            file.delete();
        }

        //Create a new file
        file.createNewFile();

        //Create a new FileWriter
        FileWriter fw=new FileWriter(file,true);

        //Create a BufferedWriter
        BufferedWriter bw=new BufferedWriter(fw);

        //Save extracted text
        bw.write(stringBuilder.toString());

        //Flush the buffer
        bw.flush();

        //Close the document
        bw.close();

        //Close the document
        fw.close();

        //Dispose the document
        doc.dispose();
    }
}
