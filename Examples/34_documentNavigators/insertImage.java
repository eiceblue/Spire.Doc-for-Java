import com.spire.doc.*;
import com.spire.doc.documents.*;

public class insertImage {
    public static void main(String[] args) throws Exception{
        String imgpath = "Data//E-iceblue.png";
        // Create a new Word document instance.
        Document doc = new Document();

        // Initialize a DocumentNavigator to help navigate and manipulate the document content.
        DocumentNavigator navigator = new DocumentNavigator(doc);

        // Write a line of text into the document indicating that an image will be inserted directly.
        navigator.writeln("Insert the picture directly:");

        // Insert an image at the current cursor position using the specified image file path.
        navigator.insertImage(imgpath);

        // Add a new section to the document.
        Section section = doc.addSection();

        // Move the cursor to the second section (index 1, since sections are zero-based).
        navigator.moveToSection(1);

        // Write a line of text explaining that the next image will have its dimensions set.
        navigator.writeln("Set the width and height of the image:");

        // Insert an image with specified width and height (both set to 100 pixels).
        navigator.insertImage(imgpath, 100, 100);

        // Write a line of text describing more advanced image positioning and formatting.
        navigator.writeln("Set the width, height, offset, and wrapping style of the image:");

        // Insert an image with detailed positioning and formatting
        navigator.insertImage(imgpath, HorizontalOrigin.Left_Margin_Area, 100, VerticalOrigin.Paragraph, 50, 100, 100, TextWrappingStyle.Through);

        // Save the document to a file.
        doc.saveToFile("InsertImage.docx", FileFormat.Docx_2019);

        // Close the document to release resources.
        doc.close();
    }
}
