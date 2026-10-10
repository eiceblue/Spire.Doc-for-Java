import com.spire.doc.*;
import com.spire.doc.documents.*;
import com.spire.doc.fields.*;

import javax.imageio.ImageIO;
import java.awt.image.*;
import java.io.*;

// Define a class to check and extract SVG images from a Word document
public class checkSVGImage {
    // Entry point of the application
    public static void main(String[] args) {
        // Create a new Document instance
        Document document = new Document();
        // Load the Word document from the specified file path
        document.loadFromFile("Data/CheckSVGImage.docx");

        // Get the body of the first section in the document
        Body body = document.getSections().get(0).getBody();
        // Declare a variable to hold the current paragraph
        Paragraph paragraph;
        // Declare a variable to hold child objects within a paragraph
        DocumentObject documentObject;
        // Loop through all paragraphs in the body
        for (int i = 0; i < body.getParagraphs().getCount(); i++) {
            // Get the paragraph at the current index
            paragraph = body.getParagraphs().get(i);
            // Loop through all child objects in the current paragraph
            for (int j = 0; j < paragraph.getChildObjects().getCount(); j++) {
                // Get the child object at the current index
                documentObject = paragraph.getChildObjects().get(j);
                // Check if the child object is a picture
                if (documentObject instanceof DocPicture) {
                    // Cast the object to DocPicture
                    DocPicture picture = (DocPicture) documentObject;
                    // Print whether the picture is in SVG format
                    System.out.println(picture.isSVG());
                    // If the picture is an SVG image
                    if (picture.isSVG()) {
                        // Get the raw byte data of the SVG image
                        byte[] svgBytes = picture.getImageBytes();
                        // Convert and save the SVG bytes to a file with a unique name
                        convert(svgBytes, "out" + i + j + ".svg");
                    }
                }
            }
        }
    }

    // Helper method to convert image bytes and save them as an SVG file
    public static void convert(byte[] imageBytes, String outSVGFile) {
        try {
            // Create an input stream from the image byte array
            ByteArrayInputStream bis = new ByteArrayInputStream(imageBytes);
            // Read the image data into a BufferedImage object
            BufferedImage image = ImageIO.read(bis);

            // Check if the image was successfully read
            if (image != null) {
                // Create a File object for the output SVG file
                File outputFile = new File(outSVGFile);
                // Write the image to the output file in SVG format
                ImageIO.write(image, "svg", outputFile);
            }
        } catch (IOException e) {
            // Print the stack trace if an I/O error occurs during file operations
            e.printStackTrace();
        }
    }
}
