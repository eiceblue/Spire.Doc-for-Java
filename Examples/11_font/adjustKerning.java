import com.spire.doc.*;
import com.spire.doc.documents.*;
import com.spire.doc.fields.*;

import java.util.*;

public class adjustKerning {
    public static void main(String[] args) {
    
        // Create a new instance of the Document class
        Document doc = new Document();

        // Add a new section to the document
        Section section = doc.addSection();

        // Create a list to store test data (text description and kerning value)
        List<Object[]> testData = new ArrayList<>();

        // Add a test case for negative kerning
        testData.add(new Object[] { "Negative Kerning (-1.0f)", -1.0f });

        // Add a test case for zero kerning (disables kerning)
        testData.add(new Object[] { "Zero Kerning (0.0f)", 0.0f });

        // Add a test case for positive kerning
        testData.add(new Object[] { "Positive Kerning (2.5f)", 2.5f });

        // Add a test case for a large kerning value
        testData.add(new Object[] { "Huge Kerning (1638.0f)", 1638.0f });

        // Add a test case for a value exceeding the standard limit (1-1638)
        testData.add(new Object[] { "Tiny Kerning (1639.0f)", 1639.0f });

        // Loop through each test data item
        for(int i=0;i<testData.size();i++) {
            // Extract the text description from the first column
            String text = (String) testData.get(i)[0];

            // Extract the kerning value from the second column
            float kerningValue = (float) testData.get(i)[1];

            // Add a new paragraph to the section
            Paragraph pragraph = section.addParagraph();

            // Append the text to the paragraph and get the text range
            TextRange textRange = pragraph.appendText(text);

            // Apply the specific kerning value to the character format
            textRange.getCharacterFormat().setKerning(kerningValue);
        }

        // Define the file name for the output document
        String result = "Adjust Kerning.docx";

        // Save the document to a file in Docx format
        doc.saveToFile(result, FileFormat.Docx);

        // Close the document to release resources
        doc.close();
        doc.dispose();
    }
}
