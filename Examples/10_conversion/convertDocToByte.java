import com.spire.doc.*;
import java.io.*;

public class convertDocToByte {
    public static void main(String[] args) {
        String input = "data/Template.docx";
        Document doc = new Document();
        // Load the document from disk.
        doc.loadFromFile(input);

        // Create a new memory stream.
        ByteArrayOutputStream outputStream = new ByteArrayOutputStream ();
        // Save the document to stream.
        doc.saveToStream(outputStream, FileFormat.Docx);

        // Convert the document to bytes.
        byte[] docBytes = outputStream.toByteArray();

        // The bytes are now ready to be stored/transmitted.

        // Now reverse the steps to load the bytes back into a document object.
        ByteArrayInputStream inputStream = new ByteArrayInputStream (docBytes);
        // Load the stream into a new document object.
        Document newDoc = new Document(inputStream);
    }
}
