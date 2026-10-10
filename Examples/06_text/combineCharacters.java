import com.spire.doc.*;
import com.spire.doc.fields.*;
import com.spire.doc.documents.*;
import com.spire.doc.formatting.*;

public class combineCharacters {
    public static void main(String[] args) {
        // Create a new Document object to hold the content
        Document doc = new Document();

        // Add a new section to the document
        Section section = doc.addSection();

        // Add a paragraph to the section for the main title
        Paragraph titlePara = section.addParagraph();

        // Append the title text for the first example
        titlePara.appendText("1. Basic Two-Lines-in-One: ");

        // Add a new paragraph for the first text content
        Paragraph p1 = section.addParagraph();

        // Append the text to be formatted with the Two-Lines-in-One effect and store the reference
        TextRange farEastLayout = p1.appendText("Basic Two-Lines-in-One: One Two Three Four");

        // Set the font size of the text range to 12 points
        farEastLayout.getCharacterFormat().setFontSize(12);

        // Set the Far East font name to Arial for the text range
        farEastLayout.getCharacterFormat().setFontNameFarEast("Arial");

        // Create a new FarEastLayout object to configure the combined character layout
        FarEastLayout layout1 = new FarEastLayout();

        // Enable the Two-Lines-in-One (combine) feature
        layout1.setCombine(true);

        // Apply the FarEastLayout settings to the text range's character format
        farEastLayout.getCharacterFormat().setFarEastLayout(layout1);

        // Add a new paragraph for the second section's title
        Paragraph p1Title = section.addParagraph();

        // Append the title text for the brackets example, preceded by two line breaks
        p1Title.appendText("\n\n2. Two-Lines-in-One + Different Brackets: ");

        // Iterate through all available bracket types in the CombineBrackets enumeration
        for (CombineBrackets bracket : CombineBrackets.values()) {

            // Add a new paragraph for the current bracket type example
            Paragraph pTmp = section.addParagraph();

            // Append text displaying the current bracket type name along with sample characters
            TextRange rt = pTmp.appendText("Bracket Type " + bracket.name() + ": A B C D");

            // Set the font size of this text range to 12 points
            rt.getCharacterFormat().setFontSize(12);

            // Set the Far East font name to Arial for this text range
            rt.getCharacterFormat().setFontNameFarEast("Arial");

            // Create a new FarEastLayout object for the current bracket configuration
            FarEastLayout layoutTmp = new FarEastLayout();

            // Enable the Two-Lines-in-One (combine) feature
            layoutTmp.setCombine(true);

            // Set the specific bracket type for the combined layout
            layoutTmp.setCombineBrackets(bracket);

            // Apply the configured FarEastLayout to the current text range
            rt.getCharacterFormat().setFarEastLayout(layoutTmp);
        }

        // Save the document to a file named "CombineCharacters.docx" in Docx 2019 format
        doc.saveToFile("CombineCharacters.docx", FileFormat.Docx_2019);

        // Close the document to release file handles
        doc.close();

        // Dispose of the document object to free up memory resources
        doc.dispose();
    }
}
