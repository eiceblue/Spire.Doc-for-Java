import com.spire.doc.*;
import java.io.*;

public class loadAndSaveToStream {
    public static void main(String[] args) throws Exception {
        String input = "data/Template.docx";
        String result = "output/loadAndSaveToStream_out.rtf";
        // Open the stream. Read only access is enough to load a document.
        InputStream stream = new FileInputStream(input);

        // Load the entire document into memory.
        Document doc = new Document(stream);

        // You can close the stream now, it is no longer needed because the document is in memory.
        stream.close();
        // Do something with the document

        // Convert the document to a different format and save to stream.
        File outFile = new File(result);
        OutputStream newStream = new FileOutputStream(outFile);
        doc.saveToStream(newStream, FileFormat.Rtf);
    }
}
