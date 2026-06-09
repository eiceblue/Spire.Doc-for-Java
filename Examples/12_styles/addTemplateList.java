import com.spire.doc.*;
import com.spire.doc.documents.*;

public class addTemplateList {
    public static void main(String[] args) {
        // Create a new Document object to represent the Word document
        Document document = new Document();

        // Add a new section to the document, which is a fundamental structural element
        Section section = document.addSection();

        // Define a default bullet list template for creating unordered lists
        ListTemplate template = ListTemplate.Bullet_Default;

        // Register the bullet list template with the document and get a reference to it
        ListDefinitionReference listRef = document.getListReferences().add(template);

        // Define a default numbered list template for creating ordered lists
        ListTemplate template1 = ListTemplate.Number_Default;

        // Register the numbered list template with the document and get a reference to it
        ListDefinitionReference listRef1 = document.getListReferences().add(template1);

        // Create a new paragraph within the current section
        Paragraph paragraph = section.addParagraph();

        // Add the text "List Item 1" to the newly created paragraph
        paragraph.appendText("List Item 1");

        // Apply the bullet list format (listRef) at level 0 to this paragraph
        paragraph.getListFormat().applyListRef(listRef, 0);

        // Create another new paragraph for the next list item
        paragraph = section.addParagraph();

        // Add the text "List Item 2" to the paragraph
        paragraph.appendText("List Item 2");

        // Apply the bullet list format at level 1 (a nested level) to this paragraph
        paragraph.getListFormat().applyListRef(listRef, 1);

        // Create a third paragraph for the final bullet point
        paragraph = section.addParagraph();

        // Add the text "List Item 3" to the paragraph
        paragraph.appendText("List Item 3");

        // Apply the bullet list format at level 2 (a deeper nested level) to this paragraph
        paragraph.getListFormat().applyListRef(listRef, 2);

        // Start the numbered list by creating a new paragraph
        paragraph = section.addParagraph();

        // Add the text "List Item 6" to the paragraph
        paragraph.appendText("List Item 6");

        // Apply the numbered list format (listRef1) at level 0 to this paragraph
        paragraph.getListFormat().applyListRef(listRef1, 0);

        // Create a new paragraph for the second numbered item
        paragraph = section.addParagraph();

        // Add the text "List Item 7" to the paragraph
        paragraph.appendText("List Item 7");

        // Apply the numbered list format at level 1 to this paragraph
        paragraph.getListFormat().applyListRef(listRef1, 1);

        // Create a new paragraph for the third numbered item
        paragraph = section.addParagraph();

        // Add the text "List Item 8" to the paragraph
        paragraph.appendText("List Item 8");

        // Apply the numbered list format at level 2 to this paragraph
        paragraph.getListFormat().applyListRef(listRef1, 2);

        // Save the completed document to a file in Docx format
        document.saveToFile("addTemplateList.docx", FileFormat.Docx);

        // Close the document, releasing any associated resources like file handles
        document.close();

        // Dispose of the document object, freeing up system memory
        document.dispose();
    }
}
