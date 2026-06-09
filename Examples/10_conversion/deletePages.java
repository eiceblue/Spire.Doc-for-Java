import com.spire.doc.Document;
import com.spire.doc.FileFormat;

import java.util.ArrayList;
import java.util.Arrays;
import java.util.List;

public class deletePages {
    public static void main(String[] args) {
    
        // Initialize a new Document object
        Document document = new Document();

        // Load an existing Word document from the specified relative file path
        document.loadFromFile("Data/RemovePages.docx");

        // Remove all blank pages from the document
        document.removeBlankPages();

        // Remove specific pages by index (0-based).
        List<Integer> list = new ArrayList<>(Arrays.asList(2, 4));
        document.removePages(list);

        // Define the output file name for the modified document
        String outputFile = "DeletePages.docx";

        // Save the document to the specified file in DOCX 2019 format
        document.saveToFile(outputFile, FileFormat.Docx_2016);

        // Close the document to release file handles
        document.close();

        // Dispose of the document object to free up memory
        document.dispose();
    }
}
