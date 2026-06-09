import com.spire.doc.*;
import com.spire.doc.documents.*;

public class addSingleLevelList {
    public static void main(String[] args) {

        // Create a new instance of the Document class to represent a Word document
        Document document = new Document();

        // Add a new section to the document, which acts as a container for content like paragraphs and tables
        Section section = document.addSection();

        // Define a list template that uses Arabic numerals (1, 2, 3) followed by a dot
        ListTemplate template = ListTemplate.Number_Arabic_Dot;

        // Register this single-level numbered list template with the document and get a reference to it
        ListDefinitionReference listRef = document.getListReferences().addSingleLevelList(template);

        // Create a new paragraph object within the current section
        Paragraph paragraph = section.addParagraph();

        // Append the text to the newly created paragraph
        paragraph.appendText("List Item 1");

        // Apply the previously defined numbered list format (listRef) at level 0 to this paragraph
        paragraph.getListFormat().applyListRef(listRef, 0);

        // Reassign the paragraph variable by adding another new paragraph to the section
        paragraph = section.addParagraph();

        // Append the text to this new paragraph
        paragraph.appendText("List Item 2");

        // Apply the same numbered list format at level 0 to continue the sequence
        paragraph.getListFormat().applyListRef(listRef, 0);

        // Create a third paragraph in the section for the next list item
        paragraph = section.addParagraph();

        // Append the text to the paragraph
        paragraph.appendText("List Item 3");

        // Apply the numbered list format at level 0 to complete the list
        paragraph.getListFormat().applyListRef(listRef, 0);

        // Save the document to a file using Docx format
        document.saveToFile("addSingleLevelList.docx", FileFormat.Docx);

        // Close the document to release system resources associated with the file
        document.close();

        // Dispose of the document object to free up memory
        document.dispose();
    }
}
