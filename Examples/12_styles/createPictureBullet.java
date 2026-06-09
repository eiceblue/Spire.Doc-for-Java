import com.spire.doc.*;
import com.spire.doc.collections.*;
import com.spire.doc.documents.*;

public class createPictureBullet {
    public static void main(String[] args) {
        // Create a new Document object to represent a Word document.
        Document document = new Document();

        // Add a new section to the document. A section is a fundamental structural unit in Word.
        Section section = document.addSection();

        // Create a bulleted list style named "bulletList".
        ListStyle listStyle = document.getStyles().add(ListType.Bulleted, "bulletList");

        // Get the collection of list levels (up to 9 levels, indexed from 0) for this list style.
        ListLevelCollection Levels = listStyle.getListRef().getLevels();

        // Enable picture bullet for level 0 (the top-level list items).
        Levels.get(0).createPictureBullet();

        // Load an image file as the bullet for level 0 using a relative path.
        Levels.get(0).getPictureBullet().loadImage("Data//Word.png");

        // Enable picture bullet for level 1 (second-level indented list items).
        Levels.get(1).createPictureBullet();

        // Load another image file as the bullet for level 1 using a relative path.
        Levels.get(1).getPictureBullet().loadImage("Data//logo.png");

        // Add a new paragraph for the first-level list item.
        Paragraph paragraph = section.addParagraph();

        // Append the text "List Item 1" to the paragraph.
        paragraph.appendText("List Item 1");

        // Apply the custom "bulletList" style to this paragraph.
        paragraph.getListFormat().applyStyle(listStyle);

        // Add another paragraph for the second-level list item.
        paragraph = section.addParagraph();

        // Append the text "List Item 1.1" to this paragraph.
        paragraph.appendText("List Item 1.1");

        // Apply the same list style to make it part of the list structure.
        paragraph.getListFormat().applyStyle(listStyle);

        // Set the list level of this paragraph to 1 (i.e., second indentation level).
        paragraph.getListFormat().setListLevelNumber(1);

        // Save the document as "CreatePictureBullet.docx" in Word 2016 (.docx) format.
        document.saveToFile("CreatePictureBullet.docx", FileFormat.Docx_2016);

        // Close the document to release internal resources such as file handles.
        document.close();

        // Explicitly dispose of the document object to ensure memory is freed immediately.
        document.dispose();
    }
}
