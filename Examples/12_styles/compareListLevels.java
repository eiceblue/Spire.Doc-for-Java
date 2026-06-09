import com.spire.doc.*;
import com.spire.doc.collections.ListLevelCollection;
import com.spire.doc.documents.*;

public class compareListLevels {
    public static void main(String[] args) {
        // Create a new instance of a Word document.
        Document document = new Document();

        // Add a new bulleted list style named "bulletList1" to the document's styles collection.
        ListStyle listStyle_1 = document.getStyles().add(ListType.Bulleted, "bulletList1");

        // Retrieve the list level collection associated with the first list style.
        ListLevelCollection Levels_1 = listStyle_1.getListRef().getLevels();

        // Get the first (top-level) list formatting object from the first style.
        ListLevel L10 = Levels_1.get(0);

        // Get the second-level list formatting object from the first style.
        ListLevel L11 = Levels_1.get(1);

        // Get the third-level list formatting object from the first style.
        ListLevel L12 = Levels_1.get(2);

        // Add another bulleted list style named "bulletList2" to the document.
        ListStyle listStyle_2 = document.getStyles().add(ListType.Bulleted, "bulletList2");

        // Retrieve the list level collection for the second list style.
        ListLevelCollection Levels_2 = listStyle_2.getListRef().getLevels();

        // Get the first-level list formatting object from the second style.
        ListLevel L20 = Levels_2.get(0);

        // Get the second-level list formatting object from the second style.
        ListLevel L21 = Levels_2.get(1);

        // Get the third-level list formatting object from the second style.
        ListLevel L22 = Levels_2.get(2);

        // Set line spacing for the first level of the first list style (1.5 times 10-point base).
        L10.getParagraphFormat().setLineSpacing(10 * 1.5f);

        // Set font size to 9 points for the second level of the first list style.
        L11.getCharacterFormat().setFontSize(9);

        // Enable legal-style numbering (e.g., 1.1, 1.2) for the second level of the first style.
        L11.isLegalStyleNumbering(true);

        // Use Arabic numerals (1, 2, 3...) as the numbering pattern.
        L11.setPatternType(ListPatternType.Arabic);

        // Specify that no character (like a dot or parenthesis) follows the number.
        L11.setFollowCharacter(FollowCharacterType.Nothing);

        // Set the bullet character using Unicode (U+006E, which is lowercase 'n'—often used as a custom bullet).
        L11.setBulletCharacter("\\x006e");

        // Align the list number to the left within its numbering space.
        L11.setNumberAlignment(ListNumberAlignment.Left);

        // Position the number 10 points to the left of the default position (negative value).
        L11.setNumberPosition(-10);

        // Set the distance between the number and the following text tab stop to 0.5 inches.
        L11.setTabSpaceAfter(0.5f);

        // Define the starting horizontal position of the actual list text as 0.5 inches.
        L11.setTextPosition(0.5f);

        // Start numbering at 4 instead of 1 for this level.
        L11.setStartAt(4);

        // Prefix each number with the word "Chapter" (e.g., "Chapter4").
        L11.setNumberPrefix("Chapter");

        // Allow numbering to restart when a higher-level item changes.
        L11.setNoRestartByHigher(false);

        // Do not inherit the numbering format from the previous (higher) level.
        L11.setUsePrevLevelPattern(false);

        // Set the font name to "Arial" for the third level of the first list style.
        L12.getCharacterFormat().setFontName("Arial");

        // Set identical line spacing for the first level of the second list style.
        L20.getParagraphFormat().setLineSpacing(10 * 1.5f);

        // Set font size to 9 points for the second level of the second list style.
        L21.getCharacterFormat().setFontSize(9);

        // Enable legal-style numbering for the second level of the second style.
        L21.isLegalStyleNumbering(true);

        // Use Arabic numerals for numbering in the second style’s second level.
        L21.setPatternType(ListPatternType.Arabic);

        // No character follows the number in the second style’s second level.
        L21.setFollowCharacter(FollowCharacterType.Nothing);

        // Use the same Unicode bullet character as in the first style.
        L21.setBulletCharacter("\\x006e");

        // Left-align the number in the second style’s second level.
        L21.setNumberAlignment(ListNumberAlignment.Left);

        // Position the number at the same horizontal offset (-10 points).
        L21.setNumberPosition(-10);

        // Set the same tab space after the number (0.5 inches).
        L21.setTabSpaceAfter(0.5f);

        // Set the same text start position (0.5 inches).
        L21.setTextPosition(0.5f);

        // Start numbering at 4 for consistency.
        L21.setStartAt(4);

        // Apply the same "Chapter" prefix.
        L21.setNumberPrefix("Chapter");

        // Allow restart by higher levels.
        L21.setNoRestartByHigher(false);

        // Do not use the previous level’s pattern.
        L21.setUsePrevLevelPattern(false);

        // Create a picture-based bullet for this level (this makes it different from L11).
        L21.createPictureBullet();

        // Set the font name to "Arial" for the third level of the second list style.
        L22.getCharacterFormat().setFontName("Arial");

        // Compare the first levels of both list styles for equality.
        Boolean r0 = L10.equals(L20);

        // Compare the second levels; note: L21 has a picture bullet while L11 does not, so they differ.
        Boolean r1 = L11.equals(L21);

        // Compare the third levels, which only differ in their parent style but have identical properties.
        Boolean r2 = L12.equals(L22);

        // Add a new section to the document to hold output paragraphs.
        Section section = document.addSection();

        // Add a paragraph describing the first comparison.
        Paragraph paragraph = section.addParagraph();

        // Set the text explaining what is being compared (level 1 vs level 1).
        paragraph.setText("Compare the first level of the first ListStyle with the first level of the second ListStyle.");

        // Add a new paragraph to display the result of the first comparison.
        paragraph = section.addParagraph();

        // Insert the boolean result (true/false) into the paragraph text.
        paragraph.setText("The comparison result is " + r0 + ".");

        // Add a paragraph describing the second comparison.
        paragraph = section.addParagraph();

        // Set the text for comparing the second levels.
        paragraph.setText("Compare the second level of the first ListStyle with the second level of the second ListStyle.");

        // Add a new paragraph for the result.
        paragraph = section.addParagraph();

        // Display the result of the second comparison (expected: false due to picture bullet).
        paragraph.setText("The comparison result is " + r1 + ".");

        // Add a paragraph describing the third comparison.
        paragraph = section.addParagraph();

        // Set the text for comparing the third levels.
        paragraph.setText("Compare the third level of the first ListStyle with the third level of the second ListStyle.");

        // Add a new paragraph for the final result.
        paragraph = section.addParagraph();

        // Display the result of the third comparison (expected: true if only font name is set identically).
        paragraph.setText("The comparison result is " + r2 + ".");

        // Save the document to a file named "result.docx" in Word 2016 format.
        document.saveToFile("compareListLevels.docx", FileFormat.Docx_2016);

        // Close the document to release internal resources.
        document.close();

        // Explicitly dispose of the document object to free memory.
        document.dispose();
    }
}
