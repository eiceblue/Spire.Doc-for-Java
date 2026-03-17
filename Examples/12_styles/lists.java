import com.spire.doc.*;
import com.spire.doc.collections.ListLevelCollection;
import com.spire.doc.documents.*;

public class lists {
    public static void main(String[] args) {
        // Set the output file path
        String output = "output/lists.docx";

        // Create a new Document object
        Document document = new Document();

        // Add a section to the document
        Section sec = document.addSection();

        // Add a paragraph for the title
        Paragraph paragraph = sec.addParagraph();
        paragraph.appendText("Lists");
        paragraph.applyStyle(BuiltinStyle.Title);

        // Add a paragraph for the heading
        paragraph = sec.addParagraph();
        paragraph.appendText("Numbered List:").getCharacterFormat().setBold(true);

        // Create a new numbered list style
        ListStyle numberList =document.getStyles().add(ListType.Numbered, "numberList");
        ListLevelCollection Levels = numberList.getListRef().getLevels();

        // Configure the levels of the numbered list style
        Levels.get(1).setNumberPrefix("\u0000.");
        Levels.get(1).setPatternType(ListPatternType.Arabic);
        Levels.get(2).setNumberPrefix("\u0000.\u0001.");
        Levels.get(2).setPatternType(ListPatternType.Arabic);

        // Create a new bulleted list style
        ListStyle bulletList =document.getStyles().add(ListType.Bulleted, "bulletList");

        // Add paragraphs with list items to the section
        paragraph = sec.addParagraph();
        paragraph.appendText("List Item 1");
        paragraph.getListFormat().applyStyle(numberList.getName());

        paragraph = sec.addParagraph();
        paragraph.appendText("List Item 2");
        paragraph.getListFormat().applyStyle(numberList.getName());

        paragraph = sec.addParagraph();
        paragraph.appendText("List Item 2.1");
        paragraph.getListFormat().applyStyle(numberList.getName());
        paragraph.getListFormat().setListLevelNumber(1);

        paragraph = sec.addParagraph();
        paragraph.appendText("List Item 2.2");
        paragraph.getListFormat().applyStyle(numberList.getName());
        paragraph.getListFormat().setListLevelNumber(1);

        paragraph = sec.addParagraph();
        paragraph.appendText("List Item 2.2.1");
        paragraph.getListFormat().applyStyle(numberList.getName());
        paragraph.getListFormat().setListLevelNumber(2);

        paragraph = sec.addParagraph();
        paragraph.appendText("List Item 2.2.2");
        paragraph.getListFormat().applyStyle(numberList.getName());
        paragraph.getListFormat().setListLevelNumber(2);

        paragraph = sec.addParagraph();
        paragraph.appendText("List Item 2.2.3");
        paragraph.getListFormat().applyStyle(numberList.getName());
        paragraph.getListFormat().setListLevelNumber(2);

        paragraph = sec.addParagraph();
        paragraph.appendText("List Item 2.3");
        paragraph.getListFormat().applyStyle(numberList.getName());
        paragraph.getListFormat().setListLevelNumber(1);

        paragraph = sec.addParagraph();
        paragraph.appendText("List Item 3");
        paragraph.getListFormat().applyStyle(numberList.getName());

        paragraph = sec.addParagraph();
        paragraph.appendText("Bulleted List:").getCharacterFormat().setBold(true);

        paragraph = sec.addParagraph();
        paragraph.appendText("List Item 1");
        paragraph.getListFormat().applyStyle(bulletList.getName());

        paragraph = sec.addParagraph();
        paragraph.appendText("List Item 2");
        paragraph.getListFormat().applyStyle(bulletList.getName());

        paragraph = sec.addParagraph();
        paragraph.appendText("List Item 2.1");
        paragraph.getListFormat().applyStyle(bulletList.getName());
        paragraph.getListFormat().setListLevelNumber(1);

        paragraph = sec.addParagraph();
        paragraph.appendText("List Item 2.2");
        paragraph.getListFormat().applyStyle(bulletList.getName());
        paragraph.getListFormat().setListLevelNumber(1);

        paragraph = sec.addParagraph();
        paragraph.appendText("List Item 3");
        paragraph.getListFormat().applyStyle(bulletList.getName());

        // Save the document to the output file in Docx format
        document.saveToFile(output, FileFormat.Docx);

        // Dispose of the document to release resources
        document.dispose();
    }
}
