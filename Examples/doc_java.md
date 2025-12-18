# Spire.Doc Hello World Example
## Create a simple Word document with "Hello World!" text
```java
//Create word document.
Document document = new Document();

//Add a new section.
Section section = document.addSection();

//Add a new paragraph.
Paragraph paragraph = section.addParagraph();

//Append Text.
paragraph.appendText("Hello World!");

//Save to file.
document.saveToFile("output/helloWorld.docx", FileFormat.Docx);

//Dispose the document
document.dispose();
```

---

# Find and Highlight Text in Word Document
## This code finds all occurrences of a specific string in a Word document and highlights them with yellow color
```java
//Find text.
TextSelection[] textSelections = document.findAllString("word", false, true);

//Loop through the textSelections
for (TextSelection selection : textSelections){
    // Set highlight.
    selection.getAsOneRange().getCharacterFormat().setHighlightColor(Color.yellow);
}
```

---

# find and highlight keywords in paragraph
## Find all occurrences of a specific keyword in a paragraph and highlight them with a color
```java
//Create word document
Document document = new Document();

//Get the first section
Section s = document.getSections().get(0);

//Get the second paragraph
Paragraph para = s.getParagraphs().get(1);

//Find all matched keywords
TextSelection[] textSelections = para.findAllString("Word", false, true);

//Highlight text
for (TextSelection selection : textSelections)
{
    selection.getAsOneRange().getCharacterFormat().setHighlightColor(new Color(255, 255, 0));
}
```

---

# Spire.Doc Document Content Replacement
## Replace content in a document with another document by finding a specific pattern
```java
//Create a regex
Pattern regex=Pattern.compile("\\[MY_DOCUMENT\\]");

//Find the text by regex
TextSelection[] textSections = document1.findAllPattern(regex);

//Travel the found pattern
for (TextSelection seletion : textSections) {
    // Get the paragraph
    Paragraph para = seletion.getAsOneRange().getOwnerParagraph();
    // Get textRange
    TextRange textRange = seletion.getAsOneRange();
    // Get the para index
    int index = section1.getBody().getChildObjects().indexOf(para);

    //Insert the paragraphs of document2
    for (Object sectionObj: document2.getSections()) {
        Section section2=(Section)sectionObj;
        for (Object paragraphObj : section2.getParagraphs()) {
            Paragraph paragraph=(Paragraph)paragraphObj;
            section1.getBody().getChildObjects().insert(index, paragraph.deepClone());
        }
    }

    // Remove the found textRange
    para.getChildObjects().remove(textRange);
}
```

---

# Spire.Doc Regex Text Replacement
## Replace text in a Word document using regular expressions
```java
// Create word document
Document doc = new Document();

// Compiles the given regular expression into a pattern.
Pattern regex = Pattern.compile("\\#\\w+\\b");

// Replace the text by regex
doc.replace(regex, "Spire.Doc");
```

---

# Spire.Doc Text Replacement with Table
## Replace specific text in a document with a table
```java
//Get the first section
Section section = document.getSections().get(0);

//Return TextSection by finding the key text string
TextSelection selection = document.findString("Christmas Day, December 25", true, true);

//Return TextRange from TextSection
TextRange range = selection.getAsOneRange();

//Get Owner-Paragraph
Paragraph paragraph = range.getOwnerParagraph();

//Get the owner TextBody
Body body = paragraph.getOwnerTextBody();

// Get the index of paragraph
int index = body.getChildObjects().indexOf(paragraph);

//Create a new table
Table table = section.addTable(true);

//Set the number of rows and columns
table.resetCells(3, 3);

// Remove the paragraph and
body.getChildObjects().remove(paragraph);

//Insert table into the collection at the specified index
body.getChildObjects().insert(index, table);
```

---

# Spire.Doc Document Replacement
## Replace text in a document with another document
```java
Document doc = new Document();
IDocument replaceDoc = new Document();

// Replace specified text with the other document
doc.replace("Document1", replaceDoc, false, true);
```

---

# Spire.Doc Find and Replace with Image
## Find text in Word document and replace with images
```java
//Find the string "E-iceblue" in the document.
TextSelection[] selections = doc.findAllString("E-iceblue", true, true);
int index = 0;
TextRange range = null;

// Remove the text and replace it with image.
for (TextSelection selection : selections) {
    //Create a DocPicture
    DocPicture pic = new DocPicture(doc);
    
    //Load an image
    pic.loadImage("data/E-iceblue.png");
    
    //Get the TextRange
    range = selection.getAsOneRange();
    
    // Get the index of the TextRange
    index = range.getOwnerParagraph().getChildObjects().indexOf(range);
    
    //Insert the picture
    range.getOwnerParagraph().getChildObjects().insert(index, pic);
    
    // Remove the text
    range.getOwnerParagraph().getChildObjects().remove(range);
}
```

---

# Spire.Doc Text Replacement
## Replace text in a Word document
```java
// Replace text
document.replace("word", "ReplacedText", false, true);
```

---

# Spire.Doc paragraph extraction
## Extract content between paragraphs from a Word document
```java
private static void ExtractBetweenParagraphs(Document sourceDocument, Document destinationDocument, int startPara, int endPara) {
    // Extract the content.
    for (int i = (startPara - 1); (i < endPara); i++) {
        // Clone the ChildObjects of source document.
        DocumentObject doobj = sourceDocument.getSections().get(0).getBody().getChildObjects().get(i).deepClone();
        
        // Add to destination document.
        destinationDocument.getSections().get(0).getBody().getChildObjects().add(doobj);
    }
}
```

---

# Spire.Doc Document Content Extraction
## Extract content between paragraphs with specific styles
```java
private static void ExtractBetweenParagraphStyles(Document sourceDocument, Document destinationDocument, String stylename1, String stylename2) {
    int startindex = 0;
    int endindex = 0;
    // travel the sections of source document
    for (Object sectionObj  : sourceDocument.getSections()) {
        Section section=(Section)sectionObj;

        // travel the paragraphs
        for (Object paragraphObj : section.getParagraphs()) {
            Paragraph paragraph=(Paragraph)paragraphObj;

            // Judge paragraph style1
            if (paragraph.getStyleName().equals(stylename1)) {
                // Get the paragraph index
                startindex = section.getBody().getParagraphs().indexOf(paragraph);
            }

            // Judge paragraph style2
            if (paragraph.getStyleName().equals(stylename2)) {
                // Get the paragraph index
                endindex = section.getBody().getParagraphs().indexOf(paragraph);
            }
        }

        // Extract the content
        for (int i = (startindex + 1); i < endindex; i++) {
            // Clone the ChildObjects of source document
            DocumentObject doobj = sourceDocument.getSections().get(0).getBody().getChildObjects().get(i).deepClone();
            // Add to destination document
            destinationDocument.getSections().get(0).getBody().getChildObjects().add(doobj);
        }
    }
}
```

---

# Extract Content from Bookmark
## Extract content from a bookmark in a Word document and add it to another document

```java
//Locate the bookmark in source document.
BookmarksNavigator navigator = new BookmarksNavigator(sourcedocument);
// Find bookmark by name.
navigator.moveToBookmark("Test", true, true);
//get text body part.
TextBodyPart textBodyPart = navigator.getBookmarkContent();

//Create a TextRange type list.
List<TextRange> list = new ArrayList<TextRange>();

// Traverse the items of text body
for (Object item : textBodyPart.getBodyItems()) {
    // if it is paragraph
    if ((item instanceof Paragraph)) {
        // Traverse the ChildObjects of the paragraph
        for (Object childObject : ((Paragraph)(item)).getChildObjects()) {
            // if it is TextRange
            if ((childObject instanceof TextRange)) {
                // Add it into list
                TextRange range = ((TextRange)(childObject));
                list.add(range);
            }
        }
    }
}

// Add the extract content to destination document
for (int m = 0; m < list.size(); m++) {
    paragraph.getItems().add(list.get(m).deepClone());
}
```

---

# Spire.Doc Comment Content Extraction
## Extract content from a comment range in a Word document
```java
//Create source and destination documents
Document sourceDoc = new Document();
Document destinationDoc = new Document();
Section destinationSec = destinationDoc.addSection();

//Get a comment from the source document
Comment comment = sourceDoc.getComments().get(0);

//Get the paragraph that contains the comment
Paragraph para = comment.getOwnerParagraph();

//Find the start and end indices of the comment range
int startIndex = para.getChildObjects().indexOf(comment.getCommentMarkStart());
int endIndex = para.getChildObjects().indexOf(comment.getCommentMarkEnd());

//Extract content within the comment range
for (int i = startIndex; (i <= endIndex); i++) {
    // Clone each element in the comment range
    DocumentObject doobj = para.getChildObjects().get(i).deepClone();
    
    // Add the cloned element to the destination document
    destinationSec.addParagraph().getChildObjects().add(doobj);
}
```

---

# Spire.Doc Document Content Extraction
## Extract content from paragraph to table
```java
private static void ExtractByTable(Document sourceDocument, Document destinationDocument, int startPara, int tableNo) {
    // Get the table from the source document
    Table table = sourceDocument.getSections().get(0).getTables().get((tableNo - 1));

    // Get the table index
    int index = sourceDocument.getSections().get(0).getBody().getChildObjects().indexOf(table);

    for (int i = (startPara - 1); (i <= index); i++) {
        // Clone the ChildObjects of source document
        DocumentObject doobj = sourceDocument.getSections().get(0).getBody().getChildObjects().get(i).deepClone();

        // Add to destination document
        destinationDocument.getSections().get(0).getBody().getChildObjects().add(doobj);
    }
}
```

---

# spire.doc extract content from form field
## Extract content starting from a form field in a Word document
```java
// Define a variable
int index = 0;

// Traverse FormFields
for (Object fieldObj : sourceDocument.getSections().get(0).getBody().getFormFields()) {
    FormField field = (FormField) fieldObj;

    // Find FieldFormTextInput type field
    if (field.getType() == FieldType.Field_Form_Text_Input) {
        // Get the paragraph
        Paragraph paragraph = field.getOwnerParagraph();
        // Get the index
        index = sourceDocument.getSections().get(0).getBody().getChildObjects().indexOf(paragraph);
        break;
    }
}

// Extract the content
for (int i = index; i < (index + 3); i++) {
    // Clone the ChildObjects of source document
    DocumentObject doobj = sourceDocument.getSections().get(0).getBody().getChildObjects().get(i).deepClone();
    // Add to destination document
    section.getBody().getChildObjects().add(doobj);
}
```

---

# Spire.Doc Section Management
## Add and delete sections in a Word document
```java
private static void AddSection(Document doc)
{
    //Add a section
    doc.addSection();
}

private static void DeleteSection(Document doc)
{
    //Delete the last section
    doc.getSections().removeAt(doc.getSections().getCount() - 1);
}
```

---

# Spire.Doc Document Section Cloning
## Clone sections from one document to another
```java
// Create the source word document
Document srcDoc = new Document();

// Create the destination word document
Document desDoc = new Document();

// Initializes a section with a null value
Section cloneSection = null;

for (Object sectionObj : srcDoc.getSections()) {
    Section section = (Section)sectionObj;
    // Clone section
    cloneSection = section.deepClone();
    // Add the cloneSection in destination file
    desDoc.getSections().add(cloneSection);
}
```

---

# Spire.Doc Section Cloning
## Clone content from one section to another in a Word document
```java
//Get the first section
Section sec1 = doc.getSections().get(0);

//Get the second section
Section sec2 = doc.getSections().get(1);

// Loop through the contents of sec1
for (Object docObj : sec1.getBody().getChildObjects()) {
    DocumentObject obj=(DocumentObject)docObj;
    // Clone the contents to sec2
    sec2.getBody().getChildObjects().add(obj.deepClone());
}
```

---

# Spire.Doc Document Section Page Setup
## Modify page margins and page size for document sections
```java
// Loop through all sections
for (Object sectionObj : doc.getSections()) {
    Section section = (Section)sectionObj;
    // Modify the margins
    section.getPageSetup().setMargins(new MarginsF(100, 80, 100, 80));

    // Modify the page size
    section.getPageSetup().setPageSize(PageSize.Letter);
}
```

---

# Remove Section Content in Word Document
## This code demonstrates how to remove content from headers, body, and footers of all sections in a Word document
```java
// Loop through all sections
for (Object sectionObj : doc.getSections()) {
    Section section = (Section)sectionObj;

    // Remove header content
    section.getHeadersFooters().getHeader().getChildObjects().clear();

    // Remove body content
    section.getBody().getChildObjects().clear();

    // Remove footer content
    section.getHeadersFooters().getFooter().getChildObjects().clear();
}
```

---

# Spire.Doc Horizontal Line
## Add a horizontal line to a Word document
```java
// Create a new Document object
Document doc = new Document();

// Add a new Section to the Document
Section sec = doc.addSection();

// Add a new Paragraph to the Section
Paragraph para = sec.addParagraph();

// Append a horizontal line to the Paragraph
para.appendHorizonalLine();
```

---

# Spire.Doc Tab Stops
## Add tab stops to paragraphs with different alignments and leader types
```java
//Add tab and set its position (in points).
Tab tab = paragraph.getFormat().getTabs().addTab(28);

// Set tab alignment.
tab.setJustification(TabJustification.Left);

// Move to next tab and append text.
paragraph.appendText("\tWashing Machine");

// Add another tab and set its position (in points).
tab = paragraph.getFormat().getTabs().addTab(280);

// Set tab alignment.
tab.setJustification(TabJustification.Left);

// Specify tab leader type.
tab.setTabLeader(TabLeader.Dotted);

// Move to next tab and append text.
paragraph.appendText("\t$650");

// Add tab with no leader
tab = paragraph.getFormat().getTabs().addTab(28);
tab.setJustification(TabJustification.Left);
paragraph.appendText("\tRefrigerator");

tab = paragraph.getFormat().getTabs().addTab(280);
tab.setJustification(TabJustification.Left);
tab.setTabLeader(TabLeader.No_Leader);
paragraph.appendText("\t$800");
```

---

# Spire.Doc Latin Text Wrap
## Allow Latin text to wrap in the middle of a word
```java
//Create Word document.
Document document = new Document();

Paragraph para = document.getSections().get(0).getParagraphs().get(0);

//Allow Latin text to wrap in the middle of a word
para.getFormat().setWordWrap(false);
```

---

# Spire.Doc Paragraph Copying
## Copy paragraphs between Word documents
```java
//Create Word document1.
Document document1 = new Document();

//Create a new document.
Document document2 = new Document();

//Get paragraph 1 and paragraph 2 in document1.
Section s = document1.getSections().get(0);
Paragraph p1 = s.getParagraphs().get(0);
Paragraph p2 = s.getParagraphs().get(1);

//Copy p1 and p2 to document2.
Section s2 = document2.addSection();
Paragraph NewPara1 = ((Paragraph)(p1.deepClone()));
s2.getParagraphs().add(NewPara1);
Paragraph NewPara2 = ((Paragraph)(p2.deepClone()));
s2.getParagraphs().add(NewPara2);
```

---

# Spire.Doc Catalogue Creation
## Create a hierarchical catalogue with custom numbered list styles in Word document
```java
//Create Word document
Document document = new Document();

//Add a new section
Section section = document.addSection();
Paragraph paragraph = section.addParagraph();

//Add Heading 1
paragraph.appendText(BuiltinStyle.Heading_1.toString());
paragraph.applyStyle(BuiltinStyle.Heading_1);
paragraph.getListFormat().applyNumberedStyle();

// Add Heading 2
paragraph = section.addParagraph();
paragraph.appendText(BuiltinStyle.Heading_2.toString());
paragraph.applyStyle(BuiltinStyle.Heading_2);

//List style for Headings 2
ListStyle listSty2 = new ListStyle(document, ListType.Numbered);
for (Object listLevelObj : listSty2.getLevels()) {
    ListLevel listLev = (ListLevel)listLevelObj;
    listLev.setUsePrevLevelPattern(true);
    listLev.setNumberPrefix("1.");
}
listSty2.setName("MyStyle2");
document.getListStyles().add(listSty2);
paragraph.getListFormat().applyStyle(listSty2.getName());

//Add list style 3
ListStyle listSty3 = new ListStyle(document, ListType.Numbered);
for (Object listLevelObj : listSty3.getLevels()) {
    ListLevel listLev = (ListLevel)listLevelObj;
    listLev.setUsePrevLevelPattern(true);
    listLev.setNumberPrefix("1.1.");
}
listSty3.setName("MyStyle3");
document.getListStyles().add(listSty3);

// Add Heading 3
for (int i = 0; i < 4; i++) {
    paragraph = section.addParagraph();
    // Append text
    paragraph.appendText(BuiltinStyle.Heading_3.toString());
    // Apply list style 3 for Heading 3
    paragraph.applyStyle(BuiltinStyle.Heading_3);
    paragraph.getListFormat().applyStyle(listSty3.getName());
}
```

---

# Spire.Doc Paragraph Style Filter
## Extract paragraphs with specific style name from Word document

```java
//Create Word document.
Document document = new Document();

// Load the file from disk.
document.loadFromFile("data/Template_Docx_3.docx");

// Get paragraphs by style name.
for (Object sectionObj : document.getSections()) {
    Section section = (Section)sectionObj;
    for (Object paragraphObj: section.getParagraphs()) {
        Paragraph paragraph = (Paragraph)paragraphObj;
        if ((paragraph.getStyleName().equals("Heading1"))) {
            //Extract text
            paragraph.getText();
        }
    }
}
```

---

# Spire.Doc Paragraph Hiding
## Hide a paragraph in a Word document by setting the Hidden property of its text ranges
```java
//Create Word document
Document document = new Document();

//Get the first section
Section sec = document.getSections().get(0);

//Get the first paragraph
Paragraph para = sec.getParagraphs().get(0);

// Loop through the textranges
for (Object docObj : para.getChildObjects()) {
    DocumentObject obj = (DocumentObject)docObj;
    if ((obj instanceof TextRange)) {
        TextRange range = ((TextRange)(obj));

        //Set CharacterFormat's Hidden property as true to hide the texts
        range.getCharacterFormat().setHidden(true);
    }
}
```

---

# Spire.Doc RTF String Insertion
## Insert RTF string into Word document
```java
//Create Word document.
Document document = new Document();

//Add a new section.
Section section = document.addSection();

//Add a paragraph to the section.
Paragraph para = section.addParagraph();

//Declare a String variable to store the Rtf string.
String rtfString = "{\\rtf1\\ansi\\deff0 {\\fonttbl {\\f0 hakuyoxingshu7000;}}\\f0\\fs28 Hello, World}";

// Append Rtf string to paragraph.
para.appendRTF(rtfString);
```

---

# Spire.Doc Paragraph Pagination Management
## Set page break before a paragraph in Word document
```java
//Get the first section
Section sec = document.getSections().get(0);

//Get the fifth paragraph
Paragraph para = sec.getParagraphs().get(4);

//Set the pagination format as PageBreakBefore for the checked paragraph
para.getFormat().setPageBreakBefore(true);
```

---

# Spire.Doc Remove All Paragraphs
## Remove all paragraphs from every section in a Word document
```java
//Remove paragraphs from every section in the document
for ( Object sectionObj: document.getSections()) {
    Section section = (Section)sectionObj;
    section.getParagraphs().clear();
}
```

---

# Word Document Empty Lines Removal
## Remove empty paragraphs from Word document sections
```java
//Traverse every section on the Word document and remove the null and empty paragraphs.
for (Object sectionObj : document.getSections()) {
    Section section=(Section)sectionObj;

    //Get child objects
    for (int i = 0; i < section.getBody().getChildObjects().getCount(); i++) {

        //Judge the type equals Paragraph or not
        if ((section.getBody().getChildObjects().get(i).getDocumentObjectType().equals(DocumentObjectType.Paragraph) )) {
           String s= ((Paragraph)(section.getBody().getChildObjects().get(i))).getText().trim();
            if (s.isEmpty()) {

                //Remove the empty paragraph
                section.getBody().getChildObjects().remove(section.getBody().getChildObjects().get(i));
                i--;
            }
        }
    }
}
```

---

# Spire.Doc Paragraph Removal
## Remove specific paragraph from Word document
```java
// Create Word document
Document document = new Document();

// Remove the first paragraph from the first section of the document
document.getSections().get(0).getParagraphs().removeAt(0);
```

---

# Spire.Doc Paragraph Spacing
## Set spacing before and after paragraph lines
```java
// Access the first section of the document
Section section = doc.getSections().get(0);

// Access the first paragraph in the section
Paragraph paragraph = section.getParagraphs().get(0);

// Set the spacing before the paragraph
paragraph.getFormat().setBeforeSpacingLines(5f);

// Set the spacing after the paragraph 
paragraph.getFormat().setAfterSpacingLines(15f);
```

---

# Spire.Doc Paragraph Indentation
## Set paragraph indentation by character count
```java
// Create a new Document object
Document document = new Document();

// Add a section to the document
Section sec = document.addSection();

// Add a paragraph for the title
Paragraph para = sec.addParagraph();
para.appendText("Paragraph Formatting");
para.applyStyle(BuiltinStyle.Title);

// Add a paragraph with indent settings
para = sec.addParagraph();
para.appendText(
        "This paragraph is indent as follows: Indent 2 characters on the left and 5 characters on the right.");
para.getFormat().setLeftIndentChars(2);
para.getFormat().setRightIndentChars(5f);
```

---

# Set Paragraph Shading in Word Document
## Demonstrates how to set background color for paragraphs and specific text in a Word document
```java
//Get a paragraph.
Paragraph paragaph = document.getSections().get(0).getParagraphs().get(0);

//Set background color for the paragraph.
paragaph.getFormat().setBackColor(Color.yellow);

//Set background color for the selected text of paragraph.
paragaph = document.getSections().get(0).getParagraphs().get(2);

//Get the target string
TextSelection selection = paragaph.find("Christmas", true, false);

//Get the text selection as a text range
TextRange range = selection.getAsOneRange();

//Get the format and set background color
range.getCharacterFormat().setTextBackgroundColor(Color.yellow);
```

---

# Spire.Doc Snap to Grid
## Set paragraph to snap to document grid
```java
//Create word document.
Document document = new Document();

//Add a new section.
Section section = document.addSection();

//define Grid pitch type
section.getPageSetup().setGridType(GridPitchType.Lines_Only);
section.getPageSetup().setLinesPerPage(15);

//Add a new paragraph.
Paragraph paragraph = section.addParagraph();

//Append Text.
paragraph.appendText("With Spire.Doc, you can generate, modify, convert, render and print documents without utilizing Microsoft Word®. But you need MS Word viewer to view the resultant document. ");

//Set snap to grid
paragraph.getFormat().setSnapToGrid(true);
```

---

# Set Paragraph Spacing
## Set spacing before and after a paragraph in a Word document
```java
//Create a new paragraph
Paragraph para = new Paragraph(document);

//Set the spacing before and after
para.getFormat().setBeforeAutoSpacing(false);
para.getFormat().setBeforeSpacing(10);
para.getFormat().setAfterAutoSpacing(false);
para.getFormat().setAfterSpacing(10);

//Insert the added paragraph to the first section
document.getSections().get(0).getParagraphs().insert(1, para);
```

---

# Spire.Doc emphasis mark application
## Apply emphasis marks to specific text in a Word document
```java
//Find text to emphasize
TextSelection[] textSelections = document.findAllString("Spire.Doc for Java", false, true);

//Set emphasis mark to the found text
for (TextSelection selection : textSelections) {
    selection.getAsOneRange().getCharacterFormat().setEmphasisMark(Emphasis.Dot);
}
```

---

# spire.doc text case conversion
## change text case to AllCaps and SmallCaps
```java
TextRange textRange;

//Get the first paragraph
Paragraph para1 = doc.getSections().get(0).getParagraphs().get(1);

//Set the text ranges' CharacterFormat to AllCaps
for (Object docObj : para1.getChildObjects()) {
    DocumentObject obj=(DocumentObject)docObj;
    if ((obj instanceof TextRange)) {
        textRange = ((TextRange)(obj));
        textRange.getCharacterFormat().setAllCaps(true);
    }
}

//Get the forth paragraph
Paragraph para2 = doc.getSections().get(0).getParagraphs().get(3);

//Set the text ranges' CharacterFormat to SmallCaps
for (Object docObj : para2.getChildObjects()) {
    DocumentObject obj=(DocumentObject)docObj;
    if ((obj instanceof TextRange)) {
        textRange = ((TextRange)(obj));
        textRange.getCharacterFormat().isSmallCaps(true);
    }
}
```

---

# Spire.Doc Barcode Creation
## Create a barcode in a Word document using Spire.Doc library
```java
//Create a document
Document doc = new Document();

//Add a paragraph
Paragraph p = doc.addSection().addParagraph();

//Add barcode and set its format
TextRange txtRang = p.appendText("H63TWX11072");
//Set barcode font name, note you need to install the barcode font on your system at first
txtRang.getCharacterFormat().setFontName("C39HrP60DlTt");

//Set the font size
txtRang.getCharacterFormat().setFontSize(80);

//Set the text color
txtRang.getCharacterFormat().setTextColor(java.awt.Color.blue);
```

---

# Spire.Doc Text Extraction
## Extract text from a Word document
```java
//Create a document
Document document = new Document();

//Load the document from disk.
document.loadFromFile("data/ExtractText.docx");

//Get text from the document
String text = document.getText();

//Dispose the document
document.dispose();
```

---

# Insert New Text in Document
## This code demonstrates how to find specific text in a document, insert new text after it, and highlight the newly added text
```java
//Find all the text string "Word" from the document
TextSelection[] selections = doc.findAllString("Word", true, true);
int index = 0;

//Define a text range
TextRange range;

//Insert new text string (New text) after the searched text string
for (TextSelection selection : selections) {
    range = selection.getAsOneRange();
    
    //Create a new TextRange
    TextRange newrange = new TextRange(doc);
    
    //Set the text
    newrange.setText("(New text)");
    
    //Get the index of the range
    index = range.getOwnerParagraph().getChildObjects().indexOf(range);
    
    //Insert the new text range after the original text range
    range.getOwnerParagraph().getChildObjects().insert((index + 1), newrange);
}

//Find and highlight the newly added text
TextSelection[] text = doc.findAllString("New text", true, true);
for (TextSelection selection : text) {
    selection.getAsOneRange().getCharacterFormat().setHighlightColor(Color.yellow);
}
```

---

# Spire.Doc Symbol Insertion
## Insert unicode symbols into a Word document
```java
//Create Word document.
Document document = new Document();

//Add a section.
Section section = document.addSection();

//Add a paragraph.
Paragraph paragraph = section.addParagraph();

//Use unicode characters to create symbol Ä.
TextRange tr = paragraph.appendText("\u00c4".toString());

//Set the color of symbol Ä.
tr.getCharacterFormat().setTextColor(Color.red);

//Add symbol Ë.
paragraph.appendText("\u00cb".toString());
```

---

# Set Superscript and Subscript in Word Document
## Code to set superscript and subscript text formatting in a Word document using Spire.Doc for Java
```java
//Create word document
Document document = new Document();

//Add a new section
Section section = document.addSection();

//Add a paragraph
Paragraph paragraph = section.addParagraph();

//Append text
paragraph.appendText("E = mc");
TextRange range1 = paragraph.appendText("2");

//Set superscript
range1.getCharacterFormat().setSubSuperScript(SubSuperScript.Super_Script);

//Append a line break
paragraph.appendBreak(BreakType.Line_Break);

//Append some text
paragraph.appendText("F");
TextRange range2 = paragraph.appendText("n");

//Set subscript
range2.getCharacterFormat().setSubSuperScript(SubSuperScript.Sub_Script);

paragraph.appendText(" = F");

//Set subscript
paragraph.appendText("n-1").getCharacterFormat().setSubSuperScript(SubSuperScript.Sub_Script);
paragraph.appendText(" + F");
paragraph.appendText("n-2").getCharacterFormat().setSubSuperScript(SubSuperScript.Sub_Script);

//Loop through the paragraph and get its child objects
for (Object i : paragraph.getItems()) {
    if (i instanceof TextRange) {
        //Set font size for text range
        ((TextRange) i).getCharacterFormat().setFontSize(36);
    }
}
```

---

# Spire.Doc Text Direction Setting
## Set text direction in Word document sections and table cells
```java
//Create a new document
Document doc = new Document();

//Add the first section
Section section1 = doc.addSection();
//Set text direction for all text in a section
section1.setTextDirection(TextDirection.Right_To_Left);

//Add the second section
Section section2 = doc.addSection();

//Add a table
Table table = section2.addTable();
table.resetCells(1, 1);
TableCell cell = table.getRows().get(0).getCells().get(0);
table.getRows().get(0).setHeight(150);
table.getRows().get(0).getCells().get(0).setWidth(10);

//Set vertical text direction of table
cell.getCellFormat().setTextDirection(TextDirection.Right_To_Left_Rotated);
cell.addParagraph().appendText("This is vertical style");
```

---

# Spire.Doc Text Splitting
## Add columns to document and set line between columns
```java
//Add a column to the first section and set width and spacing
doc.getSections().get(0).addColumn(100f, 20f);

//Add a line between the two columns
doc.getSections().get(0).getPageSetup().setColumnsLineBetween(true);
```

---

# Spire.Doc Language Dictionary
## Alter language dictionary for text in Word document
```java
//Create a Word document.
Document document = new Document();

//Add new section
Section sec = document.addSection();

//Add a paragraph to the document.
Paragraph para = sec.addParagraph();

//Add a textRange for the paragraph and append some Peru Spanish words.
TextRange txtRange = para.appendText("corrige según diccionario en inglés");
short localeId=10250;
txtRange.getCharacterFormat().setLocaleIdASCII(localeId);
```

---

# Document Format Detection
## Check the format of a Word document using Spire.Doc library

```java
//Create a document
Document doc = new Document();

//Load file
doc.loadFromFile(input);

//Get the file format
FileFormat ff = doc.getDetectedFormatType();

//Check the format info
switch (ff) {
    case Doc:
        // Microsoft Word 97-2003 document
        break;
    case Dot:
        // Microsoft Word 97-2003 template
        break;
    case Docx:
        // Office Open XML WordprocessingML Macro-Free Document
        break;
    case Docm:
        // Office Open XML WordprocessingML Macro-Enabled Document
        break;
    case Dotx:
        // Office Open XML WordprocessingML Macro-Free Template
        break;
    case Dotm:
        // Office Open XML WordprocessingML Macro-Enabled Template
        break;
    case Rtf:
        // RTF format
        break;
    case Word_ML:
        // Microsoft Word 2003 WordprocessingML format
        break;
    case Html:
        // HTML format
        break;
    case Word_Xml:
        // Microsoft Word xml format for word 2007-2013
        break;
    case Odt:
        // OpenDocument Text
        break;
    case Ott:
        // OpenDocument Text Template
        break;
    case Doc_Pre_97:
        // Microsoft Word 6 or Word 95 format
        break;
    default:
        // Unknown format
        break;
}

//Dispose the document
doc.dispose();
```

---

# Document Password Protection Check
## Check if a Word document is password protected
```java
// Determine if the document is protected or not
boolean value = Document.isEncrypted("data/decrypt.docx");
```

---

# Document Password Protection Check
## Check if a document stream is password protected
```java
// Determine if the stream is protected or not
FileInputStream inStream = new FileInputStream("data/decrypt.docx");
boolean isPwd = Document.isEncrypted(inStream);
```

---

# Document Comparison
## Compare two Word documents and highlight differences
```java
//Compare two documents
doc1.compare(doc2, "E-iceblue");
```

---

# Spire.Doc Document Comparison
## Compare documents with custom options
```java
//Create a CompareOptions
CompareOptions compareOptions = new CompareOptions();
compareOptions.setIgnoreFormatting(true);

//Compare the two documents
doc1.compare(doc2, "E-iceblue", compareOptions);
```

---

# Spire.Doc Document Comparison
## Compare documents at word level
```java
// Create a new CompareOptions object for specifying comparison options.
CompareOptions compareOptions = new CompareOptions();

// Set the comparison level to Word.
compareOptions.setTextCompareLevel(TextDiffMode.Word);

// Compare the contents of doc1 with doc2 using the specified comparison options.
doc1.compare(doc2, "E-iceblue", compareOptions);
```

---

# Document Comparison
## Compare documents ignoring headers and footers
```java
// Create documents
Document doc1 = new Document();
Document doc2 = new Document();
// Set Compare options
CompareOptions options = new CompareOptions();
options.setIgnoreHeadersAndFooters(true);
// Compare document
doc1.compare(doc2,"e-iceblue",options);
```

---

# Document Word Count
## Count words and characters in a document
```java
// Create a new Document object
Document document = new Document();

// Retrieve the character count of the document (excluding spaces)
int charCount = document.getBuiltinDocumentProperties().getCharCount();

// Retrieve the character count of the document (including spaces)
int charCountWithSpace = document.getBuiltinDocumentProperties().getCharCountWithSpace();

// Retrieve the word count of the document
int wordCount = document.getBuiltinDocumentProperties().getWordCount();
```

---

# Spire.Doc Document Properties
## Set built-in document properties for a Word document
```java
// Create a new Document object
Document document = new Document();

// Access the built-in document properties and set their values accordingly.
document.getBuiltinDocumentProperties().setTitle("Document Demo Document");
document.getBuiltinDocumentProperties().setSubject("demo");
document.getBuiltinDocumentProperties().setAuthor("James");
document.getBuiltinDocumentProperties().setCompany("e-iceblue");
document.getBuiltinDocumentProperties().setManager("Jakson");
document.getBuiltinDocumentProperties().setCategory("Doc Demos");
document.getBuiltinDocumentProperties().setKeywords("Document, Property, Demo");
document.getBuiltinDocumentProperties().setComments("This document is just a demo.");
```

---

# Spire.Doc Document Properties
## Retrieve built-in and custom document properties from a Word document
```java
// Retrieve specific built-in document properties and store their values in variables.
String title = document.getBuiltinDocumentProperties().getTitle();
String comments = document.getBuiltinDocumentProperties().getComments();
String author = document.getBuiltinDocumentProperties().getAuthor();
String keywords = document.getBuiltinDocumentProperties().getKeywords();
String company = document.getBuiltinDocumentProperties().getCompany();

// Iterate through the custom document properties and retrieve their names and values.
for (int i = 0; i < document.getCustomDocumentProperties().getCount(); i++) {
    String propertyName = document.getCustomDocumentProperties().get(i).getName();
    Object propertyValue = document.getCustomDocumentProperties().get(i).getValue();
}
```

---

# Document Loading and Saving
## Load a document from disk and save it to another location
```java
// Create a new Document object.
Document doc = new Document();

// Load the document content from the specified input file.
doc.loadFromFile(input);

// Save the loaded document to the specified output file in the Docx format.
doc.saveToFile(result, FileFormat.Docx);

// Dispose the doc object to release resources
doc.dispose();
```

---

# Spire.Doc Stream Operations
## Load document from stream and save to another stream
```java
// Define file paths
String input = "data/Template.docx";
String result = "output/loadAndSaveToStream_out.rtf";

// Create input and output streams
InputStream stream = new FileInputStream(input);
File outFile = new File(result);
OutputStream newStream = new FileOutputStream(outFile);

// Create a Document object using the input stream
Document doc = new Document(stream);

// Load the document from the input stream in Docx format
doc.loadFromStream(stream, FileFormat.Docx);
stream.close();

// Save the document to the output stream in Rtf format
doc.saveToStream(newStream, FileFormat.Rtf);

// Dispose the doc object to release resources
doc.dispose();
```

---

# Document Properties Management
## Set built-in and custom document properties in a Word document
```java
// Set built-in document properties
document.getBuiltinDocumentProperties().setTitle("Document Demo Document");
document.getBuiltinDocumentProperties().setAuthor("James");
document.getBuiltinDocumentProperties().setCompany("e-iceblue");
document.getBuiltinDocumentProperties().setKeywords("Document, Property, Demo");
document.getBuiltinDocumentProperties().setComments("This document is just a demo.");

// Get custom document properties
CustomDocumentProperties custom = document.getCustomDocumentProperties();

// Add custom properties with different value types
custom.add("e-iceblue", true);
custom.add("Authorized By", "John Smith");
```

---

# Spire.Doc Grid Properties
## Set grid properties for document sections
```java
// Set grid properties for each section in the document
for (Object sec : doc.getSections()) {
    // Set the grid type to "Lines_Only"
    ((Section)sec).getPageSetup().setGridType(GridPitchType.Lines_Only);
    
    // Set the number of lines per page to 15
    ((Section)sec).getPageSetup().setLinesPerPage(15);
}
```

---

# Spire.Doc Hyperlink Base Property
## Set hyperlink base property for Word document
```java
// Create a new Document object
Document doc = new Document();

// Set the HyperlinkBase property of the document's built-in document properties
doc.getBuiltinDocumentProperties().setHyperLinkBase("HyperLinkBaseTest");
```

---

# Spire.Doc Word Document View Settings
## Configure Word document view modes and zoom settings
```java
// Set the document view type to "Web_Layout"
document.getViewSetup().setDocumentViewType(DocumentViewType.Web_Layout);

// Set the zoom percentage to 150%
document.getViewSetup().setZoomPercent(150);

// Set the zoom type to "None"
document.getViewSetup().setZoomType(ZoomType.None);
```

---

# Spire.Doc Document Operation
## Add section from one document to another
```java
//Get the second section from source document
Section Ssection = SouDoc.getSections().get(1);

//Add the section in target document
TarDoc.getSections().add(Ssection.deepClone());
```

---

# Document Cloning with Spire.Doc
## Clone a Word document using deep clone method
```java
//Create Word document
Document document = new Document();

//Clone the Word document
Document newDoc = document.deepClone();
```

---

# Spire.Doc Document Content Copying
## Copy content from one document to another document
```java
// Create source and destination documents
Document sourceDoc = new Document();
Document destinationDoc = new Document();

// Copy content from source to destination
for (Object sectionObj : sourceDoc.getSections()) {
    Section sec = (Section)sectionObj;
    for (Object docObj : sec.getBody().getChildObjects()) {
        DocumentObject obj = (DocumentObject)docObj;
        destinationDoc.getSections().get(0).getBody().getChildObjects().add(obj.deepClone());
    }
}
```

---

# Spire.Doc Font Table Integration
## Integrate font tables from one document to another and combine documents
```java
// Create documents
Document destDoc = new Document();
Document srcDoc = new Document();

// Copy the Fonttable data from the source document to the target document
srcDoc.integrateFontTableTo(destDoc);

// Copy the sections of source document to destination document
for (Object sectionObj : srcDoc.getSections()) {
    Section section = (Section)sectionObj;
    destDoc.getSections().add(section.deepClone());
}
```

---

# Document Format Preservation
## Keep same format when appending documents
```java
// Create documents
Document srcDoc = new Document();
Document destDoc = new Document();

// Keep same format of source document
srcDoc.setKeepSameFormat(true);

// Copy the sections of source document to destination document
for (Object sectionObj : srcDoc.getSections()) {
    Section section = (Section)sectionObj;
    destDoc.getSections().add(section.deepClone());
}
```

---

# Spire.Doc Headers and Footers Linking
## Link headers and footers between documents and clone sections
```java
//Link the headers and footers in the source file
srcDoc.getSections().get(0).getHeadersFooters().getHeader().setLinkToPrevious(true);
srcDoc.getSections().get(0).getHeadersFooters().getFooter().setLinkToPrevious(true);

//Clone the sections of source to destination
for (Object sectionObj : srcDoc.getSections()) {
    Section section=(Section)sectionObj;
    dstDoc.getSections().add(section.deepClone());
}
```

---

# Spire.Doc Document Merge
## Merge two documents by appending sections from one document to another
```java
//Loop all the sections of the documentMerge
for (Object sectionObj : documentMerge.getSections()) {
    Section section=(Section)sectionObj;
    
    //Append the sections from documentMerge
    document.getSections().add(section.deepClone());
}
```

---

# Spire.Doc Document Merge
## Merge documents on the same page
```java
//Traverse sections
for (Object sectionObj : document.getSections()) {
    Section section=(Section)sectionObj;
    // Traverse body ChildObjects
    for (Object docObj : section.getBody().getChildObjects()) {
        DocumentObject obj=(DocumentObject)docObj;
        // Clone to destination document at the same page
        destinationDocument.getSections().get(0).getBody().getChildObjects().add(obj.deepClone());
    }
}
```

---

# Document Theme Preservation
## Clone styles, themes, and compatibility when appending Word documents
```java
//Create a document
Document doc = new Document();

//Create a new Word document
Document newWord = new Document();

//Clone default style, theme, compatibility from the source file to the destination document
doc.cloneDefaultStyleTo(newWord);
doc.cloneThemesTo(newWord);
doc.cloneCompatibilityTo(newWord);

//Add the cloned section to destination document
newWord.getSections().add(doc.getSections().get(0).deepClone());
```

---

# Spire.Doc Section Break Continuous
## Set section breaks to continuous in a document
```java
// Iterate through all sections
for (Object sectionObj : doc.getSections()) {
    Section section = (Section)sectionObj;
    
    // Set section break as continuous
    section.setBreakCode(SectionBreakType.No_Break);
}
```

---

# Spire.Doc Document Insertion
## Insert text from one document into another
```java
//Insert document from file
doc.insertTextFromFile("data/Template_N3.docx", FileFormat.Auto);
```

---

# Document Splitting by Page Break
## Split a Word document into multiple documents at each page break
```java
//Create Word document.
Document original = new Document();

//Create a new document
Document newWord = new Document();

//Add a new section
Section section = newWord.addSection();

//copy the default style,theme and Compatibility
original.cloneDefaultStyleTo(newWord);
original.cloneThemesTo(newWord);
original.cloneCompatibilityTo(newWord);

//Split the original Word document into separate documents according to page break.
int index = 0;

//Traverse through all sections of original document.
for (int s = 0; s < original.getSections().getCount(); s++) {
    Section sec = original.getSections().get(s);
    //Traverse through all body child objects of each section.
    for (int c = 0; c < sec.getBody().getChildObjects().getCount(); c++) {
        DocumentObject obj = sec.getBody().getChildObjects().get(c);
        if (obj instanceof Paragraph) {
            Paragraph para = (Paragraph) obj;
            sec.cloneSectionPropertiesTo(section);
            //Add paragraph object in original section into section of new document.
            section.getBody().getChildObjects().add(para.deepClone());
            for (int i = 0; i < para.getChildObjects().getCount(); i++) {
                DocumentObject parobj = para.getChildObjects().get(i);
                if (parobj instanceof Break) {
                    Break break1 = (Break) parobj;
                    if (break1.getBreakType().equals(BreakType.Page_Break)) {
                        //Get the index of page break in paragraph.
                        int indexId = para.getChildObjects().indexOf(parobj);

                        //Remove the page break from its paragraph.
                        Paragraph newPara = (Paragraph) section.getBody().getLastParagraph();
                        newPara.getChildObjects().removeAt(indexId);

                        //Create a new document and add a section.
                        newWord = new Document();
                        section = newWord.addSection();
                        original.cloneDefaultStyleTo(newWord);
                        original.cloneThemesTo(newWord);
                        original.cloneCompatibilityTo(newWord);
                        sec.cloneSectionPropertiesTo(section);
                        //Add paragraph object in original section into section of new document.
                        section.getBody().getChildObjects().add(para.deepClone());
                        if (section.getParagraphs().get(0).getChildObjects().getCount() == 0) {
                            //Remove the first blank paragraph.
                            section.getBody().getChildObjects().removeAt(0);
                        } else {
                            //Remove the child objects before the page break.
                            while (indexId >= 0) {
                                section.getParagraphs().get(0).getChildObjects().removeAt(indexId);
                                indexId--;
                            }
                        }
                    }
                }
            }
        }
        if (obj instanceof Table) {
            //Add table object in original section into section of new document.
            section.getBody().getChildObjects().add(obj.deepClone());
        }
    }
}
```

---

# Split Word Document by Section Break
## This code demonstrates how to split a Word document into multiple documents based on section breaks
```java
// Create Word document
Document document = new Document();

// Load the file from disk
document.loadFromFile("data/Template_Docx_4.docx");

// Define another new word document object
Document newWord = new Document();

// Split a Word document into multiple documents by section break
for (int i = 0; i < document.getSections().getCount(); i++){
    String result = "output/result-SplitWordFileBySectionBreak_"+i+".docx";
    newWord.getSections().add(document.getSections().get(i).deepClone());

    // Save to file
    newWord.saveToFile(result);
}

// Dispose the document
document.dispose();
newWord.dispose();
```

---

# Spire.Doc Document Splitting
## Split Word document into multiple HTML pages based on Heading1 styles
```java
private static void SplitDocIntoMultipleHtml(String input, String outDirectory)
{
    //Create a document
    Document document = new Document();
    //Load from specified path
    document.loadFromFile(input);

    Document subDoc = null;
    Boolean first = true;
    int index = 0;

    //Iterate through sections and elements
    for(int s = 0; s < document.getSections().getCount(); s++) {
        Section sec = document.getSections().get(s);
        for(int c = 0; c < sec.getBody().getChildObjects().getCount(); c++)
        {
            DocumentObject element = sec.getBody().getChildObjects().get(c);
            if (IsInNextDocument(element))
            {
                if (!first)
                {
                    //Embed css style and image data into html page
                    subDoc.getHtmlExportOptions().setCssStyleSheetType(CssStyleSheetType.Internal);
                    subDoc.getHtmlExportOptions().setImageEmbedded(true);
                    //Save to html file
                    subDoc.saveToFile(outDirectory+"/out-"+ index +".html",FileFormat.Html);
                    subDoc = null;
                }
                first = false;
            }
            if (subDoc == null)
            {
                subDoc = new Document();
                subDoc.addSection();
            }
            subDoc.getSections().get(0).getBody().getChildObjects().add(element.deepClone());
        }
    }
    if (subDoc != null)
    {
        //Embed css style and image data into html page
        subDoc.getHtmlExportOptions().setCssStyleSheetType(CssStyleSheetType.Internal);
        subDoc.getHtmlExportOptions().setImageEmbedded(true);
        index++;
        //Save to html file
        subDoc.saveToFile(outDirectory+"/out-"+ index +".html", FileFormat.Html);
        //Dispose the document
        subDoc.dispose();
    }
    //Dispose the document
    document.dispose();
}

private static Boolean IsInNextDocument(DocumentObject element)
{
    if (element instanceof Paragraph)
    {
        Paragraph p = (Paragraph)element;
        if (p.getStyleName().equals("Heading1"))
        {
            return true;
        }
    }
    return false;
}
```

---

# Spire.Doc Track Changes
## Accept or reject tracked changes in a Word document
```java
//Create Word document.
Document document = new Document();

//Get the first section and the paragraph we want to accept/reject the changes.
Section sec = document.getSections().get(0);
Paragraph para = sec.getParagraphs().get(0);

//Accept the changes or reject the changes.
para.getDocument().acceptChanges();
//para.getDocument().rejectChanges();
```

---

# Document Track Changes
## Enable track changes in a Word document
```java
//Enable the track changes.
document.setTrackChanges(true);
```

---

# Document Revision Retrieval
## Extract insert and delete revisions from a Word document
```java
public class getRevisions {
    public static void main(String[] args) throws Exception{
        Document document = new Document();
        document.loadFromFile("data/GetRevisions.docx");
        
        // Traverse sections
        for (Section sec : (Iterable<Section>) document.getSections())
        {
            // Iterate through the element under body in the section
            for(DocumentObject docItem : (Iterable<DocumentObject>)sec.getBody().getChildObjects())
            {
                if (docItem instanceof Paragraph)
                {
                    Paragraph para = (Paragraph)docItem;
                    // Check if the paragraph is an insertion revision
                    if (para.isInsertRevision())
                    {
                        // Get insertion revision information
                        EditRevision insRevison = para.getInsertRevision();
                        EditRevisionType insType = insRevison.getType();
                        String insAuthor = insRevison.getAuthor();
                    }
                    // Check if the paragraph is a delete revision
                    else if (para.isDeleteRevision())
                    {
                        // Get delete revision information
                        EditRevision delRevison = para.getDeleteRevision();
                        EditRevisionType delType = delRevison.getType();
                        String delAuthor = delRevison.getAuthor();
                    }
                    // Iterate through the elements in the paragraph
                    for(DocumentObject obj : (Iterable<DocumentObject>)para.getChildObjects())
                    {
                        if (obj instanceof TextRange)
                        {
                            TextRange textRange = (TextRange)obj;
                            // Check if the textrange is an insertion revision
                            if (textRange.isInsertRevision())
                            {
                                // Get insertion revision information
                                EditRevision insRevison = textRange.getInsertRevision();
                                EditRevisionType insType = insRevison.getType();
                                String insAuthor = insRevison.getAuthor();
                            }
                            else if (textRange.isDeleteRevision())
                            {
                                // Get delete revision information
                                EditRevision delRevison = textRange.getDeleteRevision();
                                EditRevisionType delType = delRevison.getType();
                                String delAuthor = delRevison.getAuthor();
                            }
                        }
                    }
                }
            }
        }
    }
}
```

---

# Document Revision Time Modifier
## Modify revision timestamps in a Word document
```java
// Set the desired date and time for revision changes
SimpleDateFormat formatter = new SimpleDateFormat("yyyy/MM/dd HH:mm:ss");
String dateString = "2023/3/1 00:00:00";
Date date = formatter.parse(dateString);

// Iterate through the sections of the document
for (Section sec : (Iterable<Section>) document.getSections()) {

    // Iterate through the child objects in the section's body
    for (DocumentObject docItem : (Iterable<DocumentObject>) sec.getBody().getChildObjects()) {
        if (docItem instanceof Paragraph) {
            Paragraph para = (Paragraph) docItem;

            // Check if the paragraph has an insert revision
            if (para.isInsertRevision()) {
                // Get the insert revision object and set the revision date
                EditRevision insRevison = para.getInsertRevision();
                insRevison.setDateTime(date);
            }
            // Check if the paragraph has a delete revision
            else if (para.isDeleteRevision()) {
                // Get the delete revision object and set the revision date
                EditRevision delRevison = para.getDeleteRevision();
                delRevison.setDateTime(date);
            }

            // Iterate through the child objects in the paragraph
            for (DocumentObject obj : (Iterable<DocumentObject>) para.getChildObjects()) {
                if (obj instanceof TextRange) {
                    TextRange textRange = (TextRange) obj;

                    // Check if the text range has an insert revision
                    if (textRange.isInsertRevision()) {
                        // Get the insert revision object and set the revision date
                        EditRevision insRevison = textRange.getInsertRevision();
                        insRevison.setDateTime(date);
                    }
                    // Check if the text range has a delete revision
                    else if (textRange.isDeleteRevision()) {
                        // Get the delete revision object and set the revision date
                        EditRevision delRevison = textRange.getDeleteRevision();
                        delRevison.setDateTime(date);
                    }
                }
            }
        }
    }
}
```

---

# Spire.Doc Document Variables
## Add document variables to a Word document
```java
//Create Word document.
Document document = new Document();

//Add a section.
Section section = document.addSection();

//Add a paragraph.
Paragraph paragraph = section.addParagraph();

//Add a DocVariable field.
paragraph.appendField("A1", FieldType.Field_Doc_Variable);

//Add a document variable to the DocVariable field.
document.getVariables().add("A1", "12");

//Update fields.
document.isUpdateFields(true);
```

---

# Spire.Doc Variables Counter
## Count variables in a Word document
```java
// Create Word document
Document document = new Document();

// Get the number of variables in the document
int number = document.getVariables().getCount();
```

---

# Spire.Doc Variable Removal
## Remove variables from a Word document
```java
// Create Word document
Document document = new Document();

// Remove the variable by name
document.getVariables().remove("A1");

// Update the fields
document.isUpdateFields(true);
```

---

# Spire.Doc Variables Retrieval
## Retrieve document variables by index and name
```java
//Create Word document
Document document = new Document();

//Retrieve name of the variable by index
String s1 = document.getVariables().getNameByIndex(0);

//Retrieve value of the variable by index
String s2 = document.getVariables().getValueByIndex(0);

//Retrieve the value of the variable by name
String s3 = document.getVariables().get("A1");
```

---

# Spire.Doc Gradient Background
## Set gradient background for document
```java
// Set the background type to "Gradient"
document.getBackground().setType(BackgroundType.Gradient);

// Get the gradient background object
BackgroundGradient background = document.getBackground().getGradient();

// Set the first color of the gradient background to white
background.setColor1(Color.white);

// Set the second color of the gradient background to light gray
background.setColor2(Color.lightGray);

// Set the shading variant to "Shading_Down"
background.setShadingVariant(GradientShadingVariant.Shading_Down);

// Set the shading style to "Horizontal"
background.setShadingStyle(GradientShadingStyle.Horizontal);
```

---

# Set Document Image Background
## Set an image as the background for a Word document
```java
// Set the background type to "Picture"
document.getBackground().setType(BackgroundType.Picture);

// Set the picture for the background using the specified image file path
document.getBackground().setPicture("data/Background.png");
```

---

# Spire.Doc Gutter Setup
## Add gutter to document section
```java
// Get the first section of the document
Section section = document.getSections().get(0);

// Set the gutter size to 100f (units in points)
section.getPageSetup().setGutter(100f);
```

---

# Spire.Doc Line Numbering Setup
## Configure line numbering properties for document sections
```java
// Set the start value of line numbering in the first section to 1
document.getSections().get(0).getPageSetup().setLineNumberingStartValue(1);

// Set the step value for line numbering in the first section to 6
document.getSections().get(0).getPageSetup().setLineNumberingStep(6);

// Set the distance between line numbers and text in the first section to 40f
document.getSections().get(0).getPageSetup().setLineNumberingDistanceFromText(40f);

// Set the line numbering restart mode in the first section to Continuous
document.getSections().get(0).getPageSetup().setLineNumberingRestartMode(LineNumberingRestartMode.Continuous);
```

---

# Spire.Doc Page Borders Setup
## Add page borders to a Word document
```java
// Get the first section of the document
Section section = document.getSections().get(0);

// Set the border type of the page to Double_Wave
section.getPageSetup().getBorders().setBorderType(BorderStyle.Double_Wave);

// Set the color of the page borders to light gray
section.getPageSetup().getBorders().setColor(Color.lightGray);

// Set the space (margin) on the left side of the page borders to 20
section.getPageSetup().getBorders().getLeft().setSpace(20);

// Set the space (margin) on the right side of the page borders to 20
section.getPageSetup().getBorders().getRight().setSpace(20);
```

---

# Add Page Numbers in Document Sections
## Add page numbers to footers and configure page numbering restart for sections
```java
// Iterate through the first three sections of the document
for (int i = 0; i < 3; i++) {
    // Get the footer of the current section
    HeaderFooter footer = document.getSections().get(i).getHeadersFooters().getFooter();

    // Add a paragraph to the footer
    Paragraph footerParagraph = footer.addParagraph();

    // Append a page number field to the footer paragraph
    footerParagraph.appendField("page number", FieldType.Field_Page);

    // Append " of " text to the footer paragraph
    footerParagraph.appendText(" of ");

    // Append a section pages field to the footer paragraph
    footerParagraph.appendField("number of pages", FieldType.Field_Section_Pages);

    // Set the horizontal alignment of the footer paragraph to right
    footerParagraph.getFormat().setHorizontalAlignment(HorizontalAlignment.Right);

    // If it's the last iteration, exit the loop
    if (i == 2)
        break;
    else {
        // Enable page numbering restart for the next section
        document.getSections().get(i + 1).getPageSetup().setRestartPageNumbering(true);

        // Set the starting page number for the next section to 1
        document.getSections().get(i + 1).getPageSetup().setPageStartingNumber(1);
    }
}
```

---

# Spire.Doc Different Page Setup
## Set different page setups for different sections in a Word document
```java
// Retrieve the first section of the document
Section sectionOne = doc.getSections().get(0);

// Set the orientation of the first section to Landscape
sectionOne.getPageSetup().setOrientation(PageOrientation.Landscape);

// Retrieve the second section of the document
Section sectionTwo = doc.getSections().get(1);

// Set the page size of the second section to a custom dimension of 800x800
sectionTwo.getPageSetup().setPageSize(new Dimension(800, 800));
```

---

# Insert Section Break in Document
## Core functionality for inserting section breaks in a document using Spire.Doc
```java
// Create a new Document object
Document document = new Document();

// Add a section to the document
Section section = document.addSection();

// Add a new section to the document
section = document.addSection();

// Insert a section break to start a new page after the previous section
section.addParagraph().insertSectionBreak(SectionBreakType.New_Page);

// Dispose of the Document object to release resources
document.dispose();
```

---

# Spire.Doc Page Break Insertion
## Insert page breaks after specific text in a document
```java
// Find all occurrences of the word "technology" in the document and retrieve their locations
TextSelection[] selections = document.findAllString("technology", false, true);

// Iterate over the found text selections
for (int i = 0; i < selections.length; i++) {
    TextSelection ts = selections[i];

    // Get the range of the current text selection
    TextRange range = ts.getAsOneRange();

    // Get the paragraph that contains the selected text
    Paragraph paragraph = range.getOwnerParagraph();

    // Get the index of the range within its parent paragraph
    int index = paragraph.getChildObjects().indexOf(range);

    // Create a page break object
    Break pageBreak = new Break(document, BreakType.Page_Break);

    // Insert the page break after the range within the paragraph
    paragraph.getChildObjects().insert(index + 1, pageBreak);
}
```

---

# Spire.Doc Page Break Insertion
## Insert page break in document using second approach
```java
// Get the first section of the document and access its paragraphs
Section section = document.getSections().get(0);
ParagraphCollection paragraphs = section.getParagraphs();

// Append a page break to the fourth paragraph in the section
paragraphs.get(3).appendBreak(BreakType.Page_Break);
```

---

# Spire.Doc Section Break Insertion
## Insert a section break into a Word document
```java
// Get the first section of the document and access its paragraphs
Section section = document.getSections().get(0);
ParagraphCollection paragraphs = section.getParagraphs();

// Insert a section break of type "No_Break" into the second paragraph of the section
paragraphs.get(1).insertSectionBreak(SectionBreakType.No_Break);
```

---

# Spire.Doc Page Setup
## Configure page settings, header and footer for a Word document
```java
// Set the page setup properties for the section
section.getPageSetup().setPageSize(PageSize.A4);
section.getPageSetup().getMargins().setTop(72f);
section.getPageSetup().getMargins().setBottom(72f);
section.getPageSetup().getMargins().setLeft(89.85f);
section.getPageSetup().getMargins().setRight(89.85f);

// Access the header and footer of the section
HeaderFooter header = section.getHeadersFooters().getHeader();
HeaderFooter footer = section.getHeadersFooters().getFooter();

// Add a paragraph to the header
Paragraph headerParagraph = header.addParagraph();

// Add text to the header paragraph with specific formatting
TextRange text = headerParagraph.appendText("Demo of Spire.Doc");
text.getCharacterFormat().setFontName("Arial");
text.getCharacterFormat().setFontSize(10);
text.getCharacterFormat().setItalic(true);

// Set the horizontal alignment of the header paragraph
headerParagraph.getFormat().setHorizontalAlignment(HorizontalAlignment.Right);

// Set border properties for the bottom border of the header paragraph
headerParagraph.getFormat().getBorders().getBottom().setBorderType(BorderStyle.Single);
headerParagraph.getFormat().getBorders().getBottom().setSpace(0.05F);

// Add a paragraph to the footer
Paragraph footerParagraph = footer.addParagraph();

// Add fields for page number and total number of pages in the footer
footerParagraph.appendField("page number", FieldType.Field_Page);
footerParagraph.appendText(" of ");
footerParagraph.appendField("number of pages", FieldType.Field_Num_Pages);

// Set the horizontal alignment of the footer paragraph
footerParagraph.getFormat().setHorizontalAlignment(HorizontalAlignment.Right);

// Set border properties for the top border of the footer paragraph
footerParagraph.getFormat().getBorders().getTop().setBorderType(BorderStyle.Single);
footerParagraph.getFormat().getBorders().getTop().setSpace(0.05F);
```

---

# Spire.Doc Page Break Removal
## Remove page breaks from a Word document
```java
// Iterate through the paragraphs in the first section of the document
for (int j = 0; j < document.getSections().get(0).getParagraphs().getCount(); j++) {
    // Get the current paragraph
    Paragraph p = document.getSections().get(0).getParagraphs().get(j);

    // Iterate through the child objects (elements) in the paragraph
    for (int i = 0; i < p.getChildObjects().getCount(); i++) {
        // Get the current child object
        DocumentObject obj = p.getChildObjects().get(i);

        // Check if the child object is a Break
        if (obj.getDocumentObjectType().equals(DocumentObjectType.Break)) {
            // Cast the child object to a Break
            Break b = (Break)obj;

            // Remove the Break from the collection of child objects in the paragraph
            p.getChildObjects().remove(b);
        }
    }
}
```

---

# Spire.Doc Page Number Reset
## Reset page numbering in Word document sections
```java
// Update page numbering in the footers of each section
for (int i = 0; i < document1.getSections().getCount(); i++) {
    Section sec = document1.getSections().get(i);

    // Iterate through the child objects of the footer in each section
    for (int j = 0; j < sec.getHeadersFooters().getFooter().getChildObjects().getCount(); j++) {
        DocumentObject obj = sec.getHeadersFooters().getFooter().getChildObjects().get(j);

        // Check if the child object is a Structure_Document_Tag
        if (obj.getDocumentObjectType().equals(DocumentObjectType.Structure_Document_Tag)) {
            DocumentObject para = obj.getChildObjects().get(0);

            // Iterate through the child objects of the paragraph in Structure_Document_Tag
            for (int k = 0; k < para.getChildObjects().getCount(); k++) {
                DocumentObject item = para.getChildObjects().get(k);

                // Check if the child object is a Field
                if (item.getDocumentObjectType().equals(DocumentObjectType.Field)) {

                    // Check if the Field type is Field_Num_Pages
                    if (((Field)item).getType().equals(FieldType.Field_Num_Pages)) {
                        // Change the Field type to Field_Section_Pages
                        ((Field)item).setType(FieldType.Field_Section_Pages);
                    }
                }
            }
        }
    }
}

// Set page numbering options for the second section
document1.getSections().get(1).getPageSetup().setRestartPageNumbering(true);
document1.getSections().get(1).getPageSetup().setPageStartingNumber(1);

// Set page numbering options for the third section
document1.getSections().get(2).getPageSetup().setRestartPageNumbering(true);
document1.getSections().get(2).getPageSetup().setPageStartingNumber(1);
```

---

# Set Document Gutter Position
## Demonstrates how to set gutter position and size in a document section
```java
// Create a new Document object
Document document = new Document();

// Get the first section of the document
Section section = document.getSections().get(0);

// Set the top gutter to true for the section's page setup
section.getPageSetup().isTopGutter(true);

// Set the gutter size to 100f for the section's page setup
section.getPageSetup().setGutter(100f);
```

---

# Spire.Doc Snap To Grid
## Set SnapToGrid property for paragraphs in a document
```java
// Get the first section of the document
Section section = document.getSections().get(0);

// Iterate through each paragraph in the section
for (Paragraph paragraph : (Iterable<? extends Paragraph>)section.getParagraphs()) {
    // Set the "SnapToGrid" property of the paragraph's format to true
    paragraph.getFormat().setSnapToGrid(true);
}
```

---

# Spire.Doc Document to Byte Array Conversion
## Convert a Document object to a byte array and back to a Document object
```java
// Create a new instance of the Document class
Document doc = new Document();

// Create a ByteArrayOutputStream to store the document contents
ByteArrayOutputStream outputStream = new ByteArrayOutputStream();

// Save the document to the OutputStream in Docx format
doc.saveToStream(outputStream, FileFormat.Docx);

// Get the byte array representation of the document content
byte[] docBytes = outputStream.toByteArray();

// Create a ByteArrayInputStream from the byte array
ByteArrayInputStream inputStream = new ByteArrayInputStream(docBytes);

// Create a new Document object from the ByteArrayInputStream
Document newDoc = new Document(inputStream);
```

---

# Spire.Doc Document Object to Image Conversion
## Convert various document objects (paragraphs, tables, rows, cells, and shapes) to images
```java
// Convert a paragraph to an image
private static BufferedImage ConvertParagraphToImage(Paragraph obj) {
    // Create a new document
    Document doc = new Document();

    // Add a section to the document
    Section section = doc.addSection();

    // Add a deep clone of the paragraph to the section
    section.getBody().getChildObjects().add(obj.deepClone());

    // Save the document as an image
    BufferedImage image = doc.saveToImages(0, ImageType.Bitmap);
    doc.close();

    // Return the image
    return image;
}

// Convert a table to an image
private static BufferedImage ConvertTableToImage(Table obj) {
    // Create a new document
    Document doc = new Document();

    // Add a section to the document
    Section section = doc.addSection();

    // Add a deep clone of the table to the section
    section.getBody().getChildObjects().add(obj.deepClone());

    // Save the document as an image
    BufferedImage image = doc.saveToImages(0, ImageType.Bitmap);
    doc.close();

    // Return the image
    return image;
}

// Convert a table row to an image
private static BufferedImage ConvertTableRowToImage(TableRow obj) {
    // Create a new document
    Document doc = new Document();

    // Add a section to the document
    Section section = doc.addSection();

    // Add a table to the section
    Table table = section.addTable();

    // Add a deep clone of the row to the table
    table.getRows().add(obj.deepClone());

    // Save the document as an image
    BufferedImage image = doc.saveToImages(0, ImageType.Bitmap);
    doc.close();

    // Return the image
    return image;
}

// Convert a table cell to an image
private static BufferedImage ConvertTableCellToImage(TableCell obj) {
    // Create a new document
    Document doc = new Document();

    // Add a section to the document
    Section section = doc.addSection();

    // Add a table to the section
    Table table = section.addTable();

    // Add a new row to the table and add a deep clone of the cell to it
    table.addRow().getCells().add(obj.deepClone());

    // Save the document as an image
    BufferedImage image = doc.saveToImages(0, ImageType.Bitmap);
    doc.close();

    // Return the image
    return image;
}

// Convert a shape object to an image
private static BufferedImage ConvertShapeToImage(ShapeObject obj) {
    // Create a new document
    Document doc = new Document();

    // Add a section to the document
    Section section = doc.addSection();

    // Add a paragraph to the section and add a deep clone of the shape object to it
    section.addParagraph().getChildObjects().add(obj.deepClone());

    // Save the document as an image
    BufferedImage image = doc.saveToImages(0, ImageType.Bitmap);
    doc.close();

    // Return the image
    return image;
}
```

---

# Document Format Conversion
## Convert DOC/DOCX documents to WPS format using Spire.Doc
```java
// Create a new Document object
Document document = new Document();

// Load a document from a file
document.loadFromFile("input_path.docx");

// Save the document to WPS format
document.saveToFile("output_path.wps", FileFormat.Wps);

// Dispose of the document resources
document.dispose();
```

---

# Spire.Doc Document Conversion
## Convert Word document to WPT format
```java
// Create a new Document object
Document document = new Document();

// Load a document from the specified file
document.loadFromFile("data/Sample.docx");

// Save the document to a WPT file format
document.saveToFile("output/DocToWPT.wpt", FileFormat.Wpt);

// Dispose of the document resources
document.dispose();
```

---

# HTML to Image Conversion
## Convert HTML document to image using Spire.Doc
```java
// Create a new Document object
Document document = new Document();

// Load an HTML file into the document
document.loadFromFile("input.html", FileFormat.Html, XHTMLValidationType.None);

// Convert the first page of the document to an image
BufferedImage image = document.saveToImages(0, ImageType.Bitmap);
```

---

# HTML to PDF Conversion
## Convert HTML file to PDF format using Spire.Doc library
```java
// Create a new Document object
Document document = new Document();

// Load an HTML file into the document, specifying the file format as Html and XHTML validation type as None
document.loadFromFile("data/Template_HtmlFile.html", FileFormat.Html, XHTMLValidationType.None);

// Save the document to PDF format
document.saveToFile("output/result-HtmlToPdf.pdf", FileFormat.PDF);
```

---

# HTML to XML Conversion
## Convert HTML file to XML format using Spire.Doc
```java
// Create a new Document object
Document document = new Document();

// Load an HTML file into the document
document.loadFromFile("data/Template_HtmlFile.html");

// Save the document to XML format
document.saveToFile("output/result-HtmlToXml.xml", FileFormat.Xml);

// Dispose of the document resources
document.dispose();
```

---

# HTML to XPS Conversion
## Convert HTML files to XPS format using Spire.Doc library
```java
// Create a new Document object
Document document = new Document();

// Load an HTML file into the document, specifying the file format as Html and XHTML validation type as None
document.loadFromFile("data/Template_HtmlFile.html", FileFormat.Html, XHTMLValidationType.None);

// Specify the output file path and name for the generated XPS
String result = "output/result-HtmlToXps.xps";

// Save the document to the specified file in XPS format
document.saveToFile(result, FileFormat.XPS);

// Dispose of the document resources
document.dispose();
```

---

# Spire.Doc Image to PDF Conversion
## Convert an image file to PDF format using Spire.Doc library
```java
// Create a new Document object
Document doc = new Document();

// Add a section to the document
Section section = doc.addSection();

// Add a paragraph to the section
Paragraph paragraph = section.addParagraph();

// Append the picture to the paragraph using the specified input file
DocPicture picture = paragraph.appendPicture(input);

// Set the page size of the section to A4
section.getPageSetup().setPageSize(PageSize.A4);

// Set the top margin of the page to 10f
section.getPageSetup().getMargins().setTop(10f);

// Set the left margin of the page to 25f
section.getPageSetup().getMargins().setLeft(25f);

// Save the document to the specified file in PDF format
doc.saveToFile(result, FileFormat.PDF);

// Dispose of the document resources
doc.dispose();
```

---

# Markdown to Word or PDF Conversion
## Convert Markdown files to Word (DOCX) and PDF formats using Spire.Doc library
```java
// Create a new instance of the Document class
Document document = new Document();

// Load the content of the Markdown file into the Document object
document.loadFromFile("data/MarkDownFile.md", FileFormat.Markdown);

// Save the document as a DOCX file
document.saveToFile("output/result.docx", FileFormat.Docx_2013);

// Save the document as a PDF file
document.saveToFile("output/result.pdf", FileFormat.PDF);
document.dispose();
```

---

# ODT to Word Conversion
## Convert ODT files to DOCX format using Spire.Doc
```java
// Create a new Document object
Document document = new Document();

// Load an ODT file into the document
document.loadFromFile("input.odt");

// Save the document to DOCX format
document.saveToFile("output.docx", FileFormat.Docx);

// Dispose of the document resources
document.dispose();
```

---

# RTF to PDF Conversion
## Convert RTF document to PDF format using Spire.Doc library
```java
// Create a new Document object
Document document = new Document();

// Load an RTF file into the document, specifying the file format as RTF
document.loadFromFile("input.rtf", FileFormat.Rtf);

// Save the document to PDF format
document.saveToFile("output.pdf", FileFormat.PDF);

// Dispose of the document resources
document.dispose();
```

---

# Spire.Doc document to HTML conversion
## Convert Word document to HTML format
```java
// Create word document
Document document = new Document();
document.loadFromFile(inputFile);

// Save html file
document.saveToFile(outputFile, FileFormat.Html);
```

---

# Document to Image Conversion
## Convert Word document pages to image format
```java
// Create a new Document object
Document document = new Document();

// Load the Word document
document.loadFromFile(input);

// Save the first page of the document as a BufferedImage object
BufferedImage image = document.saveToImages(0, ImageType.Bitmap);

// Write the BufferedImage object to a file in PNG format
ImageIO.write(image, "PNG", file);

// Dispose of the resources used by the Document object
document.dispose();
```

---

# Spire.Doc High-Resolution Image Conversion
## Convert Word document to high-resolution image
```java
// Create a new Document object
Document document = new Document();

// Load the Word document
document.loadFromFile(inputPath);

// Save the first page of the document as an array of BufferedImages with high resolution (300x300 pixels)
BufferedImage[] image = document.saveToImages(0, 1, ImageType.Bitmap, 300, 300);

// Write the first BufferedImage from the array to the output file in PNG format
ImageIO.write(image[0], "PNG", new File(outputPath));

// Dispose of the resources used by the Document object
document.dispose();
```

---

# Spire.Doc Document Conversion
## Convert Word document to ODT format
```java
// Create a new instance of the Document class
Document document = new Document();

// Load the Word document from the specified input file path
document.loadFromFile(input);

// Save the loaded document to the specified output file path in the ODT format
document.saveToFile(output, FileFormat.Odt);
```

---

# Spire.Doc Document to PCL Conversion
## Convert Word document to PCL format
```java
// Create a new instance of the Document class
Document document = new Document();

// Load the Word document from the specified input file
document.loadFromFile(input);

// Save the document as a PCL file to the specified output path
document.saveToFile(output, FileFormat.PCL);

// Release any system resources used by the document
document.dispose();
```

---

# Spire.Doc PDF Conversion with Author Name in Comment Labels
## Set parameter to use author names in comment labels when converting to PDF
```java
// Create a ToPdfParameterList object to set parameters for PDF conversion
ToPdfParameterList parms = new ToPdfParameterList();

// Set the option to use the author name to display comment labels to true
parms.useAuthorNameToDisplayCommentLabel(true);
```

---

# Document to PostScript Conversion
## Core functionality for converting a document to PostScript format using Spire.Doc
```java
// Create a new Document object
Document document = new Document();

// Load the document from the input file
document.loadFromFile(input);

// Save the document to the output file in PostScript format
document.saveToFile(output, FileFormat.Post_Script);

// Dispose of the document object to free up resources
document.dispose();
```

---

# Document to RTF Conversion
## Convert document to RTF format using Spire.Doc
```java
// Create a new Document object
Document document = new Document();

// Load the document from the input file
document.loadFromFile(inputFile);

// Save the document to the output file in RTF format
document.saveToFile(outputFile, FileFormat.Rtf);

// Dispose of the document object to free up resources
document.dispose();
```

---

# Document to SVG Conversion
## Convert Word document to SVG format using Spire.Doc library
```java
// Create a new Document object
Document document = new Document();

// Save the document to SVG format
document.saveToFile(outputFile, FileFormat.SVG);

// Dispose of the document object to free up resources
document.dispose();
```

---

# Spire.Doc Document to TIFF Conversion
## Convert Word document to TIFF format
```java
// Define the output file path and name for the TIFF document
String outputFile = "output/wordToTiff.tiff";

// Create a new instance of the Document class
Document doc = new Document();

// Add a new section to the document and a paragraph inside the section
Section section = doc.addSection();
Paragraph paragraph = section.addParagraph();

// Save the document to a TIFF file
doc.saveToTiff(outputFile);

// Dispose of the document object to free up resources
doc.dispose();
```

---

# Spire.Doc Document to XML Conversion
## Convert Word document to XML format
```java
// Define the input file path and name for the Word document
String inputFile = "data/toXML.doc";

// Define the output file path and name for the XML document
String outputFile = "output/toXML.xml";

// Create a new instance of the Document class
Document document = new Document();

// Load the Word document from the specified input file
document.loadFromFile(inputFile);

// Save the document to an XML file using the specified output file path and name,
// and specifying the file format as XML
document.saveToFile(outputFile, FileFormat.Xml);

// Dispose of the document object to free up resources
document.dispose();
```

---

# Spire.Doc document conversion
## Convert Word document to XPS format
```java
// Create a new Document object
Document document = new Document();

// Load the document from the specified input file
document.loadFromFile(inputFile);

// Save the document to the specified output file in XPS format
document.saveToFile(outputFile, FileFormat.XPS);

// Dispose of the resources used by the document
document.dispose();
```

---

# Spire.Doc Text to Word Conversion
## Convert text file to Word document
```java
// Create a new Document object
Document document = new Document();

// Load the content of the input text file into the document
document.loadFromFile(inputPath);

// Save the document as a Word document with the specified output path and format
document.saveToFile(outputPath, FileFormat.Docx);

// Dispose of the document resources to free up memory
document.dispose();
```

---

# Spire.Doc Word to Markdown Conversion
## Convert Word document to Markdown format using Spire.Doc library
```java
// Instantiate a new Document object
Document document = new Document();

// Load the .docx file into the Document object
document.loadFromFile("data/convertedTemplate.docx", FileFormat.Docx);

// Save the content of the Document object as a Markdown file
document.saveToFile("output/result.md", FileFormat.Markdown);

document.dispose();
```

---

# Word to Text Conversion
## Convert Word documents to text files using Spire.Doc
```java
// Create a new Document object
Document document = new Document();

// Load the content of the input Word document into the document object
document.loadFromFile(input);

// Save the document as a text file with the specified output path and format
document.saveToFile(output, FileFormat.Txt);

// Dispose of the document resources to free up memory
document.dispose();
```

---

# Spire.Doc Word to WordML/WordXML Conversion
## Convert Word documents to WordML and WordXML formats
```java
// Create a new Document object
Document document = new Document();

// Load the Word document
document.loadFromFile(inputFile);

// Save the document in WordML format
document.saveToFile(result1, FileFormat.Word_ML);

// Save the document in WordXML format
document.saveToFile(result2, FileFormat.Word_Xml);

// Clean up resources
document.dispose();
```

---

# WPS to DOC Conversion
## Convert WPS document to DOC format using Spire.Doc library
```java
// Create a new Document object
Document document = new Document();

// Load the WPS document from the specified file
document.loadFromFile("data/Sample.wps");

// Save the loaded document to DOC format
document.saveToFile("output/WPSToDoc.doc", FileFormat.Doc);

// Clean up resources and release memory
document.dispose();
```

---

# WPT to DOC Conversion
## Convert WPT format document to DOC format using Spire.Doc library
```java
// Create a new Document object
Document document = new Document();

// Load the WPT document from the specified file ("data/Sample.wpt")
document.loadFromFile("data/Sample.wpt");

// Save the loaded document to the specified output file ("output/WPTtoDoc.doc") in Doc format
document.saveToFile("output/WPTtoDoc.doc", FileFormat.Doc);

// Clean up resources and release memory used by the Document object
document.dispose();
```

---

# Spire.Doc XML to PDF Conversion
## Convert XML document to PDF format
```java
// Create a new Document object
Document document = new Document();

// Load the XML document from the specified input file
document.loadFromFile(inputFile);

// Save the loaded document to the specified output file in PDF format
document.saveToFile(outputFile, FileFormat.PDF);

// Clean up resources and release memory used by the Document object
document.dispose();
```

---

# Spire.Doc XML to Word Conversion
## Convert XML document to Word format
```java
// Create a new Document object
Document document = new Document();

// Load the XML document from the specified input file
document.loadFromFile(inputFile);

// Save the loaded document to the specified output file in DOCX format
document.saveToFile(outputFile, FileFormat.Docx);
```

---

# Spire.Doc HTML to Word Conversion
## Convert HTML file to Word document
```java
// Define the path of the input HTML file
String inputFile = "data/InputHtmlFile.html";

// Define the path of the output Word document file
String outputFile = "output/htmlFileToWord.docx";

// Create a new instance of Document
Document document = new Document();

// Load the HTML file into the document object, specifying the file format as HTML and XHTMLValidationType as None
document.loadFromFile(inputFile, FileFormat.Html, XHTMLValidationType.None);

// Save the document to the specified output file, specifying the file format as Docx
document.saveToFile(outputFile, FileFormat.Docx);

// Dispose the document object to release any resources associated with it
document.dispose();
```

---

# HTML to Word Conversion
## Convert HTML string to Word document using Spire.Doc library
```java
// Create a new document
Document document = new Document();
// Add a section to the document
Section sec = document.addSection();

// Append the HTML text to the section as a paragraph
sec.addParagraph().appendHTML(htmlText);

// Dispose of the document resources
document.dispose();
```

---

# Add Cover Image to EPUB
## This code demonstrates how to add a cover image to a document and convert it to EPUB format
```java
// Create a new Document object
Document doc = new Document();

// Load the document from the specified file path
doc.loadFromFile("data/ToEpub.doc");

// Create a DocPicture object and load the cover image
DocPicture picture = new DocPicture(doc);
picture.loadImage("data/Cover.png");

// Save the document to EPUB format with the added cover image
doc.saveToEpub("output/addCoverImage.epub", picture);

// Clean up system resources
doc.dispose();
```

---

# Document to EPUB Conversion
## Convert Word document to EPUB format
```java
// Create a new instance of the Document class
Document doc = new Document();

// Load a Word document from the specified file path
doc.loadFromFile("data/ToEpub.doc");

// Save the document to the specified file path in EPUB format
doc.saveToFile(result, FileFormat.E_Pub);

// Clean up system resources associated with the Document object
doc.dispose();
```

---

# Spire.Doc Document to HTML Conversion
## Convert Word document to HTML format
```java
// Create a new instance of the Document class
Document document = new Document();

// Load the Word document from the specified input file path
document.loadFromFile(inputFile);

// Save the document to the specified output file path in HTML format
document.saveToFile(outputFile, FileFormat.Html);

// Clean up system resources associated with the Document object
document.dispose();
```

---

# Spire.Doc HTML Export Options
## Configure HTML export options for Word document conversion
```java
Document document = new Document();

// Enable embedding images in the HTML output
document.getHtmlExportOptions().setImageEmbedded(true);

// Set the CSS style sheet type to internal
document.getHtmlExportOptions().setCssStyleSheetType(CssStyleSheetType.Internal);
```

---

# Spire.Doc HTML Conversion
## Convert document to fixed HTML format
```java
// Create a new Document object
Document document = new Document();

// Enable embedding images in the HTML output
document.getHtmlExportOptions().setImageEmbedded(true);

// Set the CSS style sheet type to internal
document.getHtmlExportOptions().setCssStyleSheetType(CssStyleSheetType.Internal);

// Embed font file
document.getHtmlExportOptions().setFontEmbedded(true);
```

---

# Disable Hyperlinks in PDF Conversion
## This code demonstrates how to disable hyperlinks when converting a Word document to PDF using Spire.Doc library
```java
// Create a ToPdfParameterList object to configure the PDF conversion options
ToPdfParameterList pdf = new ToPdfParameterList();

// Set the 'disableLink' option to true to remove the hyperlink effect in the resulting PDF page
// Set it to false to preserve the hyperlink effect
pdf.setDisableLink(true);
```

---

# Spire.Doc Font Embedding in PDF
## Convert Word document to PDF with all fonts embedded
```java
// Create a new Document object
Document document = new Document();

// Load the Word document
document.loadFromFile("input.docx");

// Create a ToPdfParameterList object to configure the PDF conversion options
ToPdfParameterList ppl = new ToPdfParameterList();

// Enable embedding of all fonts in the resulting PDF
ppl.isEmbeddedAllFonts(true);

// Save the document as a PDF file with the PDF conversion options
document.saveToFile("output.pdf", ppl);

// Dispose of the Document object to release resources
document.dispose();
```

---

# Spire.Doc PDF Conversion with Embedded Fonts
## Embed non-installed fonts when converting Word documents to PDF
```java
// Create a ToPdfParameterList object to set parameters for PDF conversion
ToPdfParameterList parms = new ToPdfParameterList();

// Create a list to store the paths of private fonts
List<PrivateFontPath> fonts = new ArrayList<PrivateFontPath>();

// Add a private font path to the list, specifying the font name and file path
fonts.add(new PrivateFontPath("PT Serif Caption", fontFile));

// Set the private font paths in the ToPdfParameterList object
parms.setPrivateFontPaths(fonts);
```

---

# Font Substitution Warning Handling
## Handles font substitution warnings during document conversion
```java
public class fontNotFoundWarning {

    public static void main(String[] args) {
        // Initialize document
        Document doc = new Document();
        // Create HandleDocumentSubstitutionWarnings object to set warning callback function
        HandleDocumentSubstitutionWarnings substitutionWarningHandler = new HandleDocumentSubstitutionWarnings();
        doc.setWarningCallback(substitutionWarningHandler);
    }

    // Create a HandleDocumentSubstitutionWarnings callback function
    static class HandleDocumentSubstitutionWarnings implements IWarningCallback
    {
        public void warning(WarningInfo info) {
            if(info.getWarningType() == WarningType.Font_Substitution)
                FontWarnings.warning(info);
        }
        public WarningInfoCollection FontWarnings = new WarningInfoCollection();
    }
}
```

---

# Spire.Doc HTML to PDF Conversion with URL Handling
## Convert HTML to PDF while handling external URL resources
```java
public class handleUrlWhileConvertHtmlToPDF {
    public static void main(String[] args) {
        // Create a new Document object
        Document document = new Document();
        
        // Set the custom URL load handler for the document
        document.HtmlUrlLoadEvent = new MyDownloadEvent();
        
        // Load the HTML file into the document
        document.loadFromFile("data/Template_HtmlFile3.html", FileFormat.Html, XHTMLValidationType.None);
        
        // Save the document as a PDF
        document.saveToFile("output/Result.pdf", FileFormat.PDF);
        
        // Dispose of the document to release resources
        document.dispose();
    }

    // Custom URL load handler class
    static class MyDownloadEvent extends HtmlUrlLoadHandler {
        @Override
        public void invoke(Object o, HtmlUrlLoadEventArgs htmlUrlLoadEventArgs) {
            try {
                // Download the bytes from the given URL
                byte[] bytes = downloadBytesFromURL(htmlUrlLoadEventArgs.getUrl());
                // Set the downloaded bytes as the data for the event
                htmlUrlLoadEventArgs.setDataBytes(bytes);
            } catch (Exception e) {
                e.printStackTrace();
            }
        }
    }

    // Method to download bytes from a URL
    public static byte[] downloadBytesFromURL(String urlString) throws Exception {
        // Create a URL object from the string
        URL url = new URL(urlString);
        // Open a connection to the URL
        HttpURLConnection connection = (HttpURLConnection) url.openConnection();
        // Set the request method to GET
        connection.setRequestMethod("GET");
        // Set connection and read timeouts
        connection.setConnectTimeout(5000);
        connection.setReadTimeout(5000);
        // Get the response code from the connection
        int responseCode = connection.getResponseCode();
        // If the response code is HTTP OK, read the content
        if (responseCode == HttpURLConnection.HTTP_OK) {
            InputStream inputStream = connection.getInputStream();
            ByteArrayOutputStream outputStream = new ByteArrayOutputStream();
            byte[] buffer = new byte[1024];
            int bytesRead;
            // Read the input stream into the output stream
            while ((bytesRead = inputStream.read(buffer)) != -1) {
                outputStream.write(buffer, 0, bytesRead);
            }
            outputStream.close();
            // Return the byte array
            return outputStream.toByteArray();
        } else {
            // Throw an exception if the response code is not HTTP OK
            throw new Exception("Failed to download content. Response code: " + responseCode);
        }
    }
}
```

---

# Spire.Doc PDF Conversion with Hidden Text
## Convert Word document to PDF while preserving hidden text
```java
// Create a new instance of the Document class
Document document = new Document();

// Load the Word document
document.loadFromFile(inputFile);

// Create an instance of the ToPdfParameterList class to specify parameters for conversion to PDF
ToPdfParameterList pdf = new ToPdfParameterList();

// Specify that hidden text should be included in the resulting PDF document
pdf.isHidden(true);

// Save the document to PDF format, using the specified conversion parameters
document.saveToFile(outputFile, pdf);

// Dispose of system resources associated with the document
document.dispose();
```

---

# Word to PDF conversion with bookmarks preservation
## Convert Word document to PDF while preserving bookmarks with custom styling
```java
public class preserveWordBookmarks {
    public static void main(String[] args) {
        // Create a new instance of the Document class
        Document document = new Document();

        // Load the Word document
        document.loadFromFile("input_path");

        // Create parameters for PDF conversion
        ToPdfParameterList toPdf = new ToPdfParameterList();

        // Set option to create bookmarks from Word bookmarks in PDF
        toPdf.setCreateWordBookmarks(true);

        // Set title for bookmarks in PDF
        toPdf.setWordBookmarksTitle("Bookmark");

        // Set color for bookmarks in PDF
        toPdf.setWordBookmarksColor(Color.GRAY);

        // Set bookmark layout event handler for customizing bookmark appearance
        document.BookmarkLayout = new BookmarkLevelHandler() {
            @Override
            public void invoke(Object sender, BookmarkLevelEventArgs args) {
                document_BookmarkLayout(sender, args);
            }
        };

        // Save document as PDF with specified parameters
        document.saveToFile("output_path", toPdf);

        // Dispose document resources
        document.dispose();
    }

    // Custom bookmark layout event handler for setting color and style of bookmarks
    private static void document_BookmarkLayout(Object sender, BookmarkLevelEventArgs args) {
        // Check bookmark level
        if (args.getBookmarkLevel().getLevel() == 2) {
            // Set color to red and style to bold for level 2 bookmarks
            args.getBookmarkLevel().setColor(Color.RED);
            args.getBookmarkLevel().setStyle(BookmarkTextStyle.Bold);
        } else if (args.getBookmarkLevel().getLevel() == 3) {
            // Set color to gray and style to italic for level 3 bookmarks
            args.getBookmarkLevel().setColor(Color.GRAY);
            args.getBookmarkLevel().setStyle(BookmarkTextStyle.Italic);
        } else {
            // Set color to green and style to regular for other bookmark levels
            args.getBookmarkLevel().setColor(Color.GREEN);
            args.getBookmarkLevel().setStyle(BookmarkTextStyle.Regular);
        }
    }
}
```

---

# Spire.Doc Custom Fonts Setting
## Set custom fonts for document conversion
```java
// Create an InputStream for the custom font file
InputStream inputStream1 = new FileInputStream("data/PT Serif Caption.ttf");

// Create an array of InputStreams containing the custom font InputStream
InputStream[] inputStreams = new InputStream[] {inputStream1};

// Set the custom fonts for the document
document.setCustomFonts(inputStreams);

// Clear the custom fonts from the document
document.clearCustomFonts();
```

---

# Spire.Doc Image Quality Setting
## Set JPEG quality for document conversion
```java
// Create a document
Document document = new Document();

// Set the JPEG image quality (0-100, where 0 is the lowest quality)
document.setJPEGQuality(40);
```

---

# Spire.Doc PDF conversion with embedded fonts
## Specify embedded fonts when converting Word document to PDF
```java
// Create a new instance of the Document class
Document document = new Document();

// Create an instance of the ToPdfParameterList class to specify parameters for conversion to PDF
ToPdfParameterList parms = new ToPdfParameterList();

// Create a list to specify the names of embedded fonts to be used in the resulting PDF document
List<String> part = new ArrayList();
part.add("PT Serif Caption");
parms.setEmbeddedFontNameList(part);

// Save the document to the specified output file in PDF format, using the specified conversion parameters
document.saveToFile(outputFile, parms);

// Dispose of system resources associated with the document
document.dispose();
```

---

# Spire.Doc Word to PDF Conversion
## Convert Word document to PDF format
```java
// Create a new instance of the Document class
Document document = new Document();

// Load the Word document from the specified input file
document.loadFromFile(inputFile);

// Save the document to the specified output file in PDF format
document.saveToFile(outputFile, FileFormat.PDF);

// Dispose of system resources associated with the document
document.dispose();
```

---

# Spire.Doc PDF Conversion with Custom Fonts
## Convert Word document to PDF using custom fonts folders
```java
// Create a new Document object
Document document = new Document();

// Load the document from the input file
document.loadFromFile(inputFile);

// When the system does not have the fonts used in a document installed, you can place the required fonts in a custom folder and then use setCustomFontsFolders to specify that the program should retrieve fonts from this path
document.setCustomFontsFolders("D:\\Fonts");

// Save the document to the output file as PDF
document.saveToFile(outputFile, FileFormat.PDF);

// Clear the custom fonts data
document.clearCustomFontsFolders();

// Clear system cached fonts
Document.clearSystemFontCache();

// Dispose the document object
document.dispose();
```

---

# Convert Word to Password Protected PDF
## Convert a Word document to PDF with password protection
```java
// Create a new instance of the Document class
Document document = new Document();

// Load the Word document
document.loadFromFile(inputFile);

// Create PDF conversion parameters with password protection
ToPdfParameterList toPdf = new ToPdfParameterList();
String password1 = "E-iceblue";
String password2 = "123";
toPdf.getPdfSecurity().encrypt(password1, password2, PdfPermissionsFlags.None, PdfEncryptionKeySize.Key_128_Bit);

// Save the document as password-protected PDF
document.saveToFile(outputFile, toPdf);

// Dispose of system resources
document.dispose();
```

---

# Spire.Doc Font Color Modification
## Change text color in Word document paragraphs
```java
// Get the first section of the document
Section section = doc.getSections().get(0);

// Get the first paragraph of the section
Paragraph p1 = section.getParagraphs().get(0);

// Loop through each child object in the first paragraph
for (int i = 0; i < p1.getChildObjects().getCount(); i++) {
    // Check if the child object is a TextRange
    if (p1.getChildObjects().get(i) instanceof TextRange) {
        // Cast the child object to TextRange
        TextRange tr = (TextRange)p1.getChildObjects().get(i);
        // Set the text color of the TextRange to red
        tr.getCharacterFormat().setTextColor(Color.red);
    }
}

// Get the second paragraph of the section
Paragraph p2 = section.getParagraphs().get(1);

// Loop through each child object in the second paragraph
for (int j = 0; j < p2.getChildObjects().getCount(); j++) {
    // Check if the child object is a TextRange
    if (p2.getChildObjects().get(j) instanceof TextRange) {
        // Cast the child object to TextRange
        TextRange tr = (TextRange)p2.getChildObjects().get(j);
        // Set the text color of the TextRange to gray
        tr.getCharacterFormat().setTextColor(Color.GRAY);
    }
}
```

---

# Spire.Doc Font Embedding
## Embed private fonts in Word document
```java
// Create a new Document object
Document doc = new Document();

// Load the document from the input file
doc.loadFromFile(inputFile);

// Get the first section of the document
Section section = doc.getSections().get(0);

// Add a new paragraph to the section
Paragraph p = section.addParagraph();

// Append text to the paragraph
TextRange range = p.appendText(
    "Spire.Doc for Java is a professional Word Java library specifically designed for developers to create, read, write, convert and print Word document files from Java platform with fast and high quality performance.");

// Set the font name to "PT Serif Caption" and font size to 20 for the text range
range.getCharacterFormat().setFontName("PT Serif Caption");
range.getCharacterFormat().setFontSize(20);

// Enable embedding fonts in the document
doc.setEmbedFontsInFile(true);

// Add the private font to the document's private font list
doc.getPrivateFontList().add(new PrivateFontPath("PT Serif Caption", fontFile));

// Save the modified document to the output file in Docx format
doc.saveToFile(outputFile, FileFormat.Docx);
```

---

# Spire.Doc Font Formatting
## Apply custom font formatting to text ranges in a document
```java
// Create a new CharacterFormat object
CharacterFormat format = new CharacterFormat(doc);

// Set the font name to Arial and font size to 16
format.setFontName("Arial");
format.setFontSize(16);

// Iterate through the child objects of the paragraph
for (int j = 0; j < p.getChildObjects().getCount(); j++) {
    // Check if the child object is a TextRange
    if (p.getChildObjects().get(j) instanceof TextRange) {
        // Convert the child object to a TextRange
        TextRange tr = (TextRange)p.getChildObjects().get(j);

        // Apply the character format to the TextRange
        tr.applyCharacterFormat(format);
    }
}
```

---

# Spire.Doc Font Fallback Rule
## Save and load font fallback rule settings for document conversion
```java
/*Instructions:
Support for switching fonts that do not support drawing characters through the FontFallbackRule method in XML when converting to a non-flow layout document.

If there is no XML available, first save an XML using saveFontFallbackRuleSettings and then manually edit the font replacement rules in the XML.
The rules consist of three attributes: Ranges correspond to Unicode ranges for each character; FallbackFonts correspond to the font names for substitution; BaseFonts correspond to the font names for characters in the document.
When editing the XML, it is important to note that the rules are searched from top to bottom for character matching.
After editing the XML, load the rules using the loadFontFallbackRuleSettings method.
*/

// Save the font fallback rule settings to an XML file
doc.saveFontFallbackRuleSettings("fontSettings.xml");

// Load the font fallback rule settings from the XML file
doc.loadFontFallbackRuleSettings("fontSettings.xml");
```

---

# ASCII Characters Bullet Style
## Create and apply bullet styles using ASCII characters in a Word document
```java
// Create a new Document object
Document document = new Document();

// Add a section to the document
Section section = document.addSection();

// Create and configure ListStyle 1
ListStyle listStyle1 = new ListStyle(document, ListType.Bulleted);
listStyle1.setName("listStyle");
listStyle1.getLevels().get(0).setBulletCharacter(ASCII2String(0x006e));
listStyle1.getLevels().get(0).getCharacterFormat().setFontName("Wingdings");
document.getListStyles().add(listStyle1);

// Create and configure ListStyle 2
ListStyle listStyle2 = new ListStyle(document, ListType.Bulleted);
listStyle2.setName("listStyle2");
listStyle2.getLevels().get(0).setBulletCharacter(ASCII2String(0x0075));
listStyle2.getLevels().get(0).getCharacterFormat().setFontName("Wingdings");
document.getListStyles().add(listStyle2);

// Create and configure ListStyle 3
ListStyle listStyle3 = new ListStyle(document, ListType.Bulleted);
listStyle3.setName("listStyle3");
listStyle3.getLevels().get(0).setBulletCharacter(ASCII2String(0x00b2));
listStyle3.getLevels().get(0).getCharacterFormat().setFontName("Wingdings");
document.getListStyles().add(listStyle3);

// Create and configure ListStyle 4
ListStyle listStyle4 = new ListStyle(document, ListType.Bulleted);
listStyle4.setName("listStyle4");
listStyle4.getLevels().get(0).setBulletCharacter(ASCII2String(0x00d8));
listStyle4.getLevels().get(0).getCharacterFormat().setFontName("Wingdings");
document.getListStyles().add(listStyle4);

// Add Paragraph 1 to the section and apply ListStyle 1
Paragraph p1 = section.getBody().addParagraph();
p1.appendText("Spire.Doc for Java");
p1.getListFormat().applyStyle(listStyle1.getName());

// Add Paragraph 2 to the section and apply ListStyle 2
Paragraph p2 = section.getBody().addParagraph();
p2.appendText("Spire.Doc for Java");
p2.getListFormat().applyStyle(listStyle2.getName());

// Add Paragraph 3 to the section and apply ListStyle 3
Paragraph p3 = section.getBody().addParagraph();
p3.appendText("Spire.Doc for Java");
p3.getListFormat().applyStyle(listStyle3.getName());

// Add Paragraph 4 to the section and apply ListStyle 4
Paragraph p4 = section.getBody().addParagraph();
p4.appendText("Spire.Doc for Java");
p4.getListFormat().applyStyle(listStyle4.getName());

// Method to convert ASCII value to a string
public static String ASCII2String(int ascii) {
    return String.valueOf((char) ascii);
}
```

---

# Spire.Doc Character Formatting
## Demonstrates various character formatting options in a Word document
```java
// Create a new Document object
Document document = new Document();

// Add a section to the document
Section sec = document.addSection();

// Add a title paragraph with "Font Styles and Effects" text and apply the Title style
Paragraph titleParagraph = sec.addParagraph();
titleParagraph.appendText("Font Styles and Effects ");
titleParagraph.applyStyle(BuiltinStyle.Title);

// Add a regular paragraph for each character formatting example

// Strikethrough Text
Paragraph paragraph = sec.addParagraph();
TextRange tr = paragraph.appendText("Strikethough Text");
tr.getCharacterFormat().isStrikeout(true);

// Shadow Text
paragraph.appendBreak(BreakType.Line_Break);
tr = paragraph.appendText("Shadow Text");
tr.getCharacterFormat().isShadow(true);

// Small caps Text
paragraph.appendBreak(BreakType.Line_Break);
tr = paragraph.appendText("Small caps Text");
tr.getCharacterFormat().isSmallCaps(true);

// Double Strikethough Text
paragraph.appendBreak(BreakType.Line_Break);
tr = paragraph.appendText("Double Strikethough Text");
tr.getCharacterFormat().setDoubleStrike(true);

// Outline Text
paragraph.appendBreak(BreakType.Line_Break);
tr = paragraph.appendText("Outline Text");
tr.getCharacterFormat().isOutLine(true);

// AllCaps Text
paragraph.appendBreak(BreakType.Line_Break);
tr = paragraph.appendText("AllCaps Text");
tr.getCharacterFormat().setAllCaps(true);

// SubScript and SuperScript Text
paragraph.appendBreak(BreakType.Line_Break);
tr = paragraph.appendText("Text");
tr = paragraph.appendText("SubScript");
tr.getCharacterFormat().setSubSuperScript(SubSuperScript.Sub_Script);
tr = paragraph.appendText("And");
tr = paragraph.appendText("SuperScript");
tr.getCharacterFormat().setSubSuperScript(SubSuperScript.Super_Script);

// Emboss Text
paragraph.appendBreak(BreakType.Line_Break);
tr = paragraph.appendText("Emboss Text");
tr.getCharacterFormat().setEmboss(true);
tr.getCharacterFormat().setTextColor(Color.white);

// Hidden Text
paragraph.appendBreak(BreakType.Line_Break);
tr = paragraph.appendText("Hidden:");
tr = paragraph.appendText("Hidden Text");
tr.getCharacterFormat().setHidden(true);

// Engrave Text
paragraph.appendBreak(BreakType.Line_Break);
tr = paragraph.appendText("Engrave Text");
tr.getCharacterFormat().setEngrave(true);
tr.getCharacterFormat().setTextColor(Color.white);

// WesternFonts and Chinese fonts
paragraph.appendBreak(BreakType.Line_Break);
tr = paragraph.appendText("WesternFonts 中文字体");
tr.getCharacterFormat().setFontNameAscii("Calibri");
tr.getCharacterFormat().setFontNameNonFarEast("Calibri");
tr.getCharacterFormat().setFontNameFarEast("Simsun-ExtB");

// Font Size
paragraph.appendBreak(BreakType.Line_Break);
tr = paragraph.appendText("Font Size");
tr.getCharacterFormat().setFontSize(20);

// Font Color
paragraph.appendBreak(BreakType.Line_Break);
tr = paragraph.appendText("Font Color");
tr.getCharacterFormat().setTextColor(Color.red);

// Bold Italic Text
paragraph.appendBreak(BreakType.Line_Break);
tr = paragraph.appendText("Bold Italic Text");
tr.getCharacterFormat().setBold(true);
tr.getCharacterFormat().setItalic(true);

// Underline Style
paragraph.appendBreak(BreakType.Line_Break);
tr = paragraph.appendText("Underline Style");
tr.getCharacterFormat().setUnderlineStyle(UnderlineStyle.Single);

// Highlight Text
paragraph.appendBreak(BreakType.Line_Break);
tr = paragraph.appendText("Highlight Text");
tr.getCharacterFormat().setHighlightColor(Color.yellow);

// Text has shading
paragraph.appendBreak(BreakType.Line_Break);
tr = paragraph.appendText("Text has shading");
tr.getCharacterFormat().setTextBackgroundColor(Color.GREEN);

// Border Around Text
paragraph.appendBreak(BreakType.Line_Break);
tr = paragraph.appendText("Border Around Text");
tr.getCharacterFormat().getBorder().setBorderType(BorderStyle.Single);

// Text Scale
paragraph.appendBreak(BreakType.Line_Break);
tr = paragraph.appendText("Text Scale");
tr.getCharacterFormat().setTextScale((short)150);

// Character Spacing is 2 point
paragraph.appendBreak(BreakType.Line_Break);
tr = paragraph.appendText("Character Spacing is 2 point");
tr.getCharacterFormat().setCharacterSpacing(2);
```

---

# Spire.Doc Document Style Copying
## Copy styles from one Word document to another using Spire.Doc for Java
```java
// Create a new Document object for the source document
Document srcDoc = new Document();

// Create a new Document object for the destination document
Document destDoc = new Document();

// Get the StyleCollection from the source document
StyleCollection styles = srcDoc.getStyles();

// Copy each style from the source document to the destination document
for (int i = 0; i < styles.getCount(); i++) {
    destDoc.getStyles().add(styles.get(i));
}
```

---

# Spire.Doc Character Spacing
## Extract character spacing from document text ranges
```java
// Get the first section of the document
Section section = document.getSections().get(0);

// Get the first paragraph of the section
Paragraph p = section.getParagraphs().get(0);

// Iterate through the child objects of the paragraph
for (int j = 0; j < p.getChildObjects().getCount(); j++) {
    // Check if the child object is a TextRange
    if (p.getChildObjects().get(j) instanceof TextRange) {
        // Cast the child object to a TextRange
        TextRange tr = (TextRange)p.getChildObjects().get(j);

        // Get the font name and character spacing of the TextRange
        String fontName = tr.getCharacterFormat().getFontName();
        float fontSpacing = tr.getCharacterFormat().getCharacterSpacing();
    }
}
```

---

# Spire.Doc Text Extraction by Style
## Extract text from a Word document based on paragraph style name
```java
// Initialize an empty string to store the extracted text
String text = "";

// Iterate through the sections of the document
for (int i = 0; i < doc.getSections().getCount(); i++) {
    // Get the current section
    Section section = doc.getSections().get(i);

    // Iterate through the paragraphs in the section
    for (int j = 0; j < section.getParagraphs().getCount(); j++) {
        // Get the current paragraph
        Paragraph para = section.getParagraphs().get(j);

        // Get the style name of the paragraph
        String name = para.getStyleName();

        // Check if the paragraph has the desired style (Heading1)
        if (para.getStyleName().equals("Heading1")) {
            // Append the text of the paragraph to the result string
            text += para.getText();
        }
    }
}
```

---

# Spire.Doc List Styles
## Create numbered and bulleted lists in Word document
```java
// Create a new Document object
Document document = new Document();

// Add a section to the document
Section sec = document.addSection();

// Create a new numbered list style
ListStyle numberList = new ListStyle(document, ListType.Numbered);
numberList.setName("numberList");

// Configure the levels of the numbered list style
numberList.getLevels().get(1).setNumberPrefix("\u0000.");
numberList.getLevels().get(1).setPatternType(ListPatternType.Arabic);
numberList.getLevels().get(2).setNumberPrefix("\u0000.\u0001.");
numberList.getLevels().get(2).setPatternType(ListPatternType.Arabic);

// Create a new bulleted list style
ListStyle bulletList = new ListStyle(document, ListType.Bulleted);
bulletList.setName("bulletList");

// Add the list styles to the document
document.getListStyles().add(numberList);
document.getListStyles().add(bulletList);

// Add numbered list items
Paragraph paragraph = sec.addParagraph();
paragraph.appendText("List Item 1");
paragraph.getListFormat().applyStyle(numberList.getName());

paragraph = sec.addParagraph();
paragraph.appendText("List Item 2");
paragraph.getListFormat().applyStyle(numberList.getName());

// Add sub-level numbered list items
paragraph = sec.addParagraph();
paragraph.appendText("List Item 2.1");
paragraph.getListFormat().applyStyle(numberList.getName());
paragraph.getListFormat().setListLevelNumber(1);

paragraph = sec.addParagraph();
paragraph.appendText("List Item 2.2");
paragraph.getListFormat().applyStyle(numberList.getName());
paragraph.getListFormat().setListLevelNumber(1);

// Add deeper level numbered list items
paragraph = sec.addParagraph();
paragraph.appendText("List Item 2.2.1");
paragraph.getListFormat().applyStyle(numberList.getName());
paragraph.getListFormat().setListLevelNumber(2);

// Add bulleted list items
paragraph = sec.addParagraph();
paragraph.appendText("List Item 1");
paragraph.getListFormat().applyStyle(bulletList.getName());

paragraph = sec.addParagraph();
paragraph.appendText("List Item 2");
paragraph.getListFormat().applyStyle(bulletList.getName());

// Add sub-level bulleted list items
paragraph = sec.addParagraph();
paragraph.appendText("List Item 2.1");
paragraph.getListFormat().applyStyle(bulletList.getName());
paragraph.getListFormat().setListLevelNumber(1);
```

---

# Spire.Doc Multiple Text Styles
## Apply different styles to text ranges within a paragraph
```java
// Create a new Document object
Document doc = new Document();

// Add a section to the document
Section section = doc.addSection();

// Add a paragraph to the section
Paragraph para = section.addParagraph();

// Append text with multiple styles to the paragraph
TextRange range = para.appendText("Spire.Doc for Java ");
range.getCharacterFormat().setFontName("Calibri");
range.getCharacterFormat().setFontSize(16);
range.getCharacterFormat().setTextColor(Color.blue);
range.getCharacterFormat().setBold(true);
range.getCharacterFormat().setUnderlineStyle(UnderlineStyle.Single);

range = para.appendText("is a professional Word Java library");
range.getCharacterFormat().setFontName("Calibri");
range.getCharacterFormat().setFontSize(15);
```

---

# Spire.Doc Paragraph Formatting
## Demonstrates various paragraph formatting options in a Word document
```java
// Create a new Document object
Document document = new Document();

// Add a section to the document
Section sec = document.addSection();

// Add a paragraph for the title
Paragraph para = sec.addParagraph();
para.appendText("Paragraph Formatting");
para.applyStyle(BuiltinStyle.Title);

// Add a paragraph with borders
para = sec.addParagraph();
para.appendText("This paragraph is surrounded with borders.");
para.getFormat().getBorders().setBorderType(BorderStyle.Single);
para.getFormat().getBorders().setColor(Color.red);

// Add paragraphs with different horizontal alignments
para = sec.addParagraph();
para.appendText("The alignment of this paragraph is Left.");
para.getFormat().setHorizontalAlignment(HorizontalAlignment.Left);

para = sec.addParagraph();
para.appendText("The alignment of this paragraph is Center.");
para.getFormat().setHorizontalAlignment(HorizontalAlignment.Center);

para = sec.addParagraph();
para.appendText("The alignment of this paragraph is Right.");
para.getFormat().setHorizontalAlignment(HorizontalAlignment.Right);

para = sec.addParagraph();
para.appendText("The alignment of this paragraph is justified.");
para.getFormat().setHorizontalAlignment(HorizontalAlignment.Justify);

para = sec.addParagraph();
para.appendText("The alignment of this paragraph is distributed.");
para.getFormat().setHorizontalAlignment(HorizontalAlignment.Distribute);

// Add a paragraph with a gray shadow background color
para = sec.addParagraph();
para.appendText("This paragraph has the gray shadow.");
para.getFormat().setBackColor(Color.gray);

// Add a paragraph with indentations
para = sec.addParagraph();
para.appendText(
    "This paragraph has the following indentations: Left indentation is 10pt, right indentation is 10pt, first line indentation is 15pt.");
para.getFormat().setLeftIndent(10);
para.getFormat().setRightIndent(10);
para.getFormat().setFirstLineIndent(15);

// Add a paragraph with hanging indentation
para = sec.addParagraph();
para.appendText("The hanging indentation of this paragraph is 15pt.");
para.getFormat().setFirstLineIndent(-15);

// Add a paragraph with spacing settings
para = sec.addParagraph();
para.appendText(
    "This paragraph has the following spacing: spacing before is 10pt, spacing after is 20pt, line spacing is at least 10pt.");
para.getFormat().setAfterSpacing(20);
para.getFormat().setBeforeSpacing(10);
para.getFormat().setLineSpacingRule(LineSpacingRule.At_Least);
para.getFormat().setLineSpacing(10);
```

---

# Retrieve Document Styles
## Extract style names from paragraphs in a Word document
```java
// Initialize a variable to store style names
String StyleName = "";

// Iterate through sections and paragraphs to retrieve style names
for (int i = 0; i < doc.getSections().getCount(); i++) {
    Section section = doc.getSections().get(i);
    for (int j = 0; j < section.getParagraphs().getCount(); j++) {
        Paragraph para = section.getParagraphs().get(j);
        StyleName += para.getStyleName();
    }
}
```

---

# Document Styles Management
## Create and apply various document styles in Word document
```java
// Create a new Document instance
Document document = new Document();

// Add a new section to the document
Section sec = document.addSection();

// Define a title style
Style titleStyle = document.addStyle(BuiltinStyle.Title);
titleStyle.getCharacterFormat().setFontName("cambria");
titleStyle.getCharacterFormat().setFontSize(28f);
titleStyle.getCharacterFormat().setTextColor(new Color(42, 123, 136));

// Customize the bottom border of the title style
if (titleStyle instanceof ParagraphStyle) {
    ParagraphStyle ps = (ParagraphStyle)titleStyle;
    ps.getParagraphFormat().getBorders().getBottom().setBorderType(BorderStyle.Single);
    ps.getParagraphFormat().getBorders().getBottom().setColor(new Color(42, 123, 136));
    ps.getParagraphFormat().getBorders().getBottom().setLineWidth(1.5f);
    ps.getParagraphFormat().setHorizontalAlignment(HorizontalAlignment.Left);
}

// Define a normal style
Style normalStyle = document.addStyle(BuiltinStyle.Normal);
normalStyle.getCharacterFormat().setFontName("cambria");
normalStyle.getCharacterFormat().setFontSize(11f);

// Define a heading 1 style
Style heading1Style = document.addStyle(BuiltinStyle.Heading_1);
heading1Style.getCharacterFormat().setFontName("cambria");
heading1Style.getCharacterFormat().setFontSize(14f);
heading1Style.getCharacterFormat().setTextColor(new Color(42, 123, 136));

// Define a heading 2 style
Style heading2Style = document.addStyle(BuiltinStyle.Heading_2);
heading2Style.getCharacterFormat().setFontName("cambria");
heading2Style.getCharacterFormat().setFontSize(12f);
heading2Style.getCharacterFormat().setBold(true);

// Define a bullet list style
ListStyle bulletList = new ListStyle(document, ListType.Bulleted);
bulletList.getCharacterFormat().setFontName("cambria");
bulletList.getCharacterFormat().setFontSize(12f);
bulletList.setName("bulletList");
document.getListStyles().add(bulletList);

// Create a paragraph and apply the title style to it
Paragraph paragraph = sec.addParagraph();
paragraph.applyStyle(BuiltinStyle.Title);

// Create a paragraph and apply the normal style to it
paragraph = sec.addParagraph();
paragraph.applyStyle(BuiltinStyle.Normal);

// Create a paragraph and apply the heading 1 style to it
paragraph = sec.addParagraph();
paragraph.applyStyle(BuiltinStyle.Heading_1);

// Create a new paragraph and add it to the section
paragraph = sec.addParagraph();
paragraph.applyStyle(BuiltinStyle.Normal);

// Create a new paragraph and add it to the section
paragraph = sec.addParagraph();
paragraph.applyStyle(BuiltinStyle.Heading_1);

// Create a new paragraph and add it to the section
paragraph = sec.addParagraph();
paragraph.applyStyle(BuiltinStyle.Heading_2);

// Create a new paragraph and add it to the section
paragraph = sec.addParagraph();
paragraph.getListFormat().applyStyle("bulletList");
```

---

# Adding Hyperlinks to Merged Images
## This code demonstrates how to add hyperlinks to images during mail merge in a Word document
```java
public class addHyperlinkToMergedImage {
    public static void main(String[] args) throws Exception {
        // Create a new Document object
        Document spireDoc = new Document();

        // Define field names and values for mail merge
        String[] fieldNames = new String[]{"ImageFile"};
        String[] fieldValues = new String[]{"path/to/image.png"};

        // Set up the MergeImageField event handler for custom processing
        spireDoc.getMailMerge().MergeImageField = new MergeImageFieldEventHandler() {
            @Override
            public void invoke(Object sender, MergeImageFieldEventArgs args) {
                mailMerge_MergeImageField(sender, args);
            }
        };

        // Execute the mail merge with the specified field names and values
        spireDoc.getMailMerge().execute(fieldNames, fieldValues);
    }

    // Custom mail merge event handler for image fields
    private static void mailMerge_MergeImageField(Object sender, MergeImageFieldEventArgs field) {
        // Get the file path of the image
        String filePath = field.getImageFileName();

        // Check if the file path is not null and not empty
        if (filePath != null && !"".equals(filePath)) {
            try {
                // Set the image using the file path
                field.setImage(filePath);
                // Adding hyperlink for image
                field.setImageLink("https://www.baidu.com/");
            } catch (Exception e) {
                // Handle any exceptions that occur during image merging
                e.printStackTrace();
            }
        }
    }
}
```

---

# Mail Merge Alternate Row Colors
## Implementing alternating row colors in mail merge operations
```java
// Global variable to keep track of the row index
static int rowIndex = 0;

// Set up a merge field event handler
doc.getMailMerge().MergeField = new MergeFieldEventHandler() {
    public void invoke(Object sender, MergeFieldEventArgs args) {
        mailMerge_MergeField(sender, args);
    }
};

// Event handler for the merge field event
private static void mailMerge_MergeField(Object sender, MergeFieldEventArgs args) {
    // Check if the current merge field name is "Name"
    if (args.getCurrentMergeField().getFieldName().equals("Name")) {
        Color rowColor;
        // Alternate the row color based on the row index
        if (rowIndex % 2 == 0) {
            rowColor = Color.gray;
        }
        else {
            rowColor = Color.lightGray;
        }

        // Get the cell and row containing the current merge field
        TableCell cell = (TableCell) args.getCurrentMergeField().getOwnerParagraph().getOwner();
        TableRow row = cell.getOwnerRow();

        // Set the background color of the row
        row.getRowFormat().setBackColor(rowColor);

        // Increment the row index
        rowIndex++;
    }
}
```

---

# Spire.Doc Conditional Field Execution
## Create and execute conditional IF fields in a Word document using mail merge
```java
// Create a new Document object
Document doc = new Document();

// Add a section to the document
Section section = doc.addSection();

// Add a paragraph to the section and create the first IF field
Paragraph paragraph = section.addParagraph();
CreateIFField1(doc, paragraph);

// Add another paragraph to the section and create the second IF field
paragraph = section.addParagraph();
CreateIFField2(doc, paragraph);

// Define field names and values for mail merge
String[] fieldName = { "Count", "Age" };
String[] fieldValue = { "2", "30" };

// Execute the mail merge with the specified field names and values
doc.getMailMerge().execute(fieldName, fieldValue);

// Set the document to update fields
doc.isUpdateFields(true);

// Method to create the first IF field
private static void CreateIFField1(Document document, Paragraph paragraph) {
    // Create a new IfField object and set its type and code
    IfField ifField = new IfField(document);
    ifField.setType(FieldType.Field_If);
    ifField.setCode("IF ");

    // Add the IfField to the paragraph
    paragraph.getItems().add(ifField);

    // Append the merge field, comparison operator, and text for condition results
    paragraph.appendField("Count", FieldType.Field_Merge_Field);
    paragraph.appendText(" > ");
    paragraph.appendText("\"1\" ");
    paragraph.appendText("\"Greater than one\" ");
    paragraph.appendText("\"Less than one\"");

    // Create a FieldMark object for the end of the IfField
    FieldMark fieldMark = new FieldMark(document);
    fieldMark.setType(FieldMarkType.Field_End);

    // Add the FieldMark to the paragraph and set it as the end of the IfField
    paragraph.getItems().add(fieldMark);
    ifField.setEnd(fieldMark);
}

// Method to create the second IF field
private static void CreateIFField2(Document document, Paragraph paragraph) {
    // Create a new IfField object and set its type and code
    IfField ifField = new IfField(document);
    ifField.setType(FieldType.Field_If);
    ifField.setCode("IF ");

    // Add the IfField to the paragraph
    paragraph.getItems().add(ifField);

    // Append the merge field, comparison operator, and text for condition results
    paragraph.appendField("Age", FieldType.Field_Merge_Field);
    paragraph.appendText(" > ");
    paragraph.appendText("\"50\" ");
    paragraph.appendText("\"The old man\" ");
    paragraph.appendText("\"The young man\"");

    // Create a FieldMark object for the end of the IfField
    FieldMark fieldMark = new FieldMark(document);
    fieldMark.setType(FieldMarkType.Field_End);

    // Add the FieldMark to the paragraph and set it as the end of the IfField
    paragraph.getItems().add(fieldMark);
    ifField.setEnd(fieldMark);
}
```

---

# Spire.Doc Mail Merge with Region
## Execute mail merge with region functionality
```java
// Create a new Document object
Document doc = new Document();

// Execute mail merge with region using the specified data
doc.getMailMerge().executeWidthRegion(data);
```

---

# Spire.Doc Mail Merge Hide Empty Regions
## Hide empty regions during mail merge in Word documents

```java
// Create a new Document object
Document document = new Document();

// Get the current date and format it as a string
Date currentTime = new Date();
SimpleDateFormat formatter = new SimpleDateFormat("yyyy-MM-dd HH:mm:ss");
String dateString = formatter.format(currentTime);

// Define field names and values for mail merge
String[] fieldNames = new String[] {"Contact Name", "Fax", "Date"};
String[] fieldValues = new String[] {"John Smith", "+1 (69) 123456", dateString};

// Set the hide empty group option to true
document.getMailMerge().setHideEmptyGroup(true);

// Set the hide empty paragraphs option to true
document.getMailMerge().setHideEmptyParagraphs(true);

// Execute the mail merge with the specified field names and values
document.getMailMerge().execute(fieldNames, fieldValues);
```

---

# Spire.Doc Mail Merge Field Identification
## Identify merge field names and groups in a Word document
```java
// Create a new Document object
Document document = new Document();

// Get the merge group names
String[] GroupNames = document.getMailMerge().getMergeGroupNames();

// Get the merge field names within a specific group ("Products")
String[] MergeFieldNamesWithinRegion = document.getMailMerge().getMergeFieldNames("Products");

// Get all of the merge field names in the document
String[] MergeFieldNames = document.getMailMerge().getMergeFieldNames();
```

---

# Spire.Doc Mail Merge
## Execute mail merge operation with field names and values
```java
// Create a new Document object
Document document = new Document();

// Define field names and values for mail merge
String[] fieldNames = new String[] {"Contact Name", "Fax", "Date"};
String[] fieldValues = new String[] {"John Smith", "+1 (69) 123456", "2023-01-01 12:00:00"};

// Execute the mail merge with the specified field names and values
document.getMailMerge().execute(fieldNames, fieldValues);
```

---

# spire.doc mail merge form fields
## Mail merge with custom form field processing
```java
// Set up the MergeField event handler for custom processing
doc.getMailMerge().MergeField = new MergeFieldEventHandler() {
    public void invoke(Object sender, MergeFieldEventArgs args) {
        mailMerge_MergeField(sender, args);
    }
};

// Execute the mail merge with the specified field names and values
doc.getMailMerge().execute(fieldNames, fieldValues);

// Custom mail merge event handler
private static void mailMerge_MergeField(Object sender, MergeFieldEventArgs args) {
    // Check if the field value is "Yes"
    if (args.getFieldValue().equals("Yes")) {
        // Process the "Yes" field value
        String checkBoxName = args.getFieldName();
        Paragraph para = args.getCurrentMergeField().getOwnerParagraph();
        int index = para.getChildObjects().indexOf(args.getCurrentMergeField());
        CheckBoxFormField field = (CheckBoxFormField) para.appendField(checkBoxName, FieldType.Field_Form_Check_Box);
        para.getChildObjects().insert(index, field);
        para.getChildObjects().remove(args.getCurrentMergeField());
        field.setChecked(true);
    }
    // Check if the field value is "No"
    if (args.getFieldValue().equals("No")) {
        // Process the "No" field value
        String checkBoxName = args.getFieldName();
        Paragraph para = args.getCurrentMergeField().getOwnerParagraph();
        int index = para.getChildObjects().indexOf(args.getCurrentMergeField());
        CheckBoxFormField field = (CheckBoxFormField) para.appendField(checkBoxName, FieldType.Field_Form_Check_Box);
        para.getChildObjects().insert(index, field);
        para.getChildObjects().remove(args.getCurrentMergeField());
        field.setChecked(false);
    }
    // Check if the field name is "Body"
    if (args.getFieldName().equals("Body")) {
        // Process the "Body" field
        Paragraph para = args.getCurrentMergeField().getOwnerParagraph();
        para.appendHTML(args.getFieldValue().toString());
        para.getChildObjects().remove(args.getCurrentMergeField());
    }
    // Check if the field name is "Date"
    if (args.getFieldName().equals("Date")) {
        // Process the "Date" field
        String textInputName = args.getFieldName();
        Paragraph para = args.getCurrentMergeField().getOwnerParagraph();
        TextFormField field = (TextFormField) para.appendField(textInputName, FieldType.Field_Form_Text_Input);
        para.getChildObjects().remove(args.getCurrentMergeField());
        field.setText(args.getFieldValue().toString());
    }
}
```

---

# Spire.Doc Mail Merge with Images
## Demonstrates how to perform a mail merge operation with image fields in a Word document

```java
public class mailMergeImage {
    public static void main(String[] args) throws Exception {
        Document spireDoc = new Document();

        // Define field names and values for mail merge
        String[] fieldNames = new String[]{"ImageFile"};
        String[] fieldValues = new String[]{"path/to/image.png"};

        // Set up the MergeImageField event handler for custom processing
        spireDoc.getMailMerge().MergeImageField = new MergeImageFieldEventHandler() {
            @Override
            public void invoke(Object sender, MergeImageFieldEventArgs args) {
                mailMerge_MergeImageField(sender, args);
            }
        };

        // Execute the mail merge with the specified field names and values
        spireDoc.getMailMerge().execute(fieldNames, fieldValues);
    }

    // Custom mail merge event handler for image fields
    private static void mailMerge_MergeImageField(Object sender, MergeImageFieldEventArgs field) {
        // Get the file path of the image
        String filePath = field.getImageFileName();

        // Check if the file path is not null and not empty
        if (filePath != null && !"".equals(filePath)) {
            try {
                // Set the image using the file path
                field.setImage(filePath);
            } catch (Exception e) {
                // Handle any exceptions that occur during image merging
                e.printStackTrace();
            }
        }
    }
}
```

---

# spire.doc mail merge
## execute mail merge with field names and values
```java
// Create a new Document object
Document document = new Document();

// Define field names and values for mail merge
String[] fieldName = new String[] {"XX_Name"};
String[] fieldValue = new String[] {"Jason Tang"};

// Execute the mail merge with the specified field names and values
document.getMailMerge().execute(fieldName, fieldValue);
```

---

# Mail Merge Event Handler
## Custom handling of mail merge events with page breaks between groups
```java
public class mergeEventHandler {
    // Index of the last processed row
    private static int lastIndex;

    // Set up the MergeField event handler for custom processing
    public static void setupMailMergeEventHandler(Document document) {
        // Initialize the lastIndex variable
        lastIndex = 0;
        
        document.getMailMerge().MergeField = new MergeFieldEventHandler() {
            @Override
            public void invoke(Object sender, MergeFieldEventArgs args) {
                // Check if the row index is greater than the lastIndex
                if (args.getRowIndex() > lastIndex) {
                    // Update the lastIndex
                    lastIndex = args.getRowIndex();
                    // Add a page break for the merge field
                    addPageBreakForMergeField(args.getCurrentMergeField());
                }
            }
        };
    }

    // Method to add a page break for a merge field
    private static void addPageBreakForMergeField(IMergeField mergeField) {
        boolean foundGroupStart = false;
        Paragraph paramgraph = (Paragraph) mergeField.getPreviousSibling().getOwner();

        while (!foundGroupStart) {
            paramgraph = (Paragraph) paramgraph.getPreviousSibling();
            for (int i = 0; i < paramgraph.getItems().getCount(); i++) {
                ParagraphBase paraBase = paramgraph.getItems().get(i);

                if (paraBase instanceof MergeField) {
                    MergeField merageField = (MergeField) paraBase;
                    // Check if the merge field is a GroupStart field
                    if ((merageField != null) && ("GroupStart".equals(merageField.getPrefix()))) {
                        foundGroupStart = true;
                        break;
                    }
                }
            }
        }

        // Append a page break to the paragraph
        paramgraph.appendBreak(BreakType.Page_Break);
    }
}
```

---

# Spire.Doc Nested Mail Merge
## Execute nested mail merge operation with document and XML data
```java
// Create and load document
Document document = new Document();
document.loadFromFile("input_path");

// Define nested relationships for mail merge
List list = new ArrayList();
Map<String, String> dictionaryEntry = new HashMap<>();
dictionaryEntry.put("Customer", "");
list.add(dictionaryEntry.entrySet().iterator().next());

dictionaryEntry = new HashMap<>();
dictionaryEntry.put("Order", "Customer_Id = %Customer.Customer_Id%");
list.add(dictionaryEntry.entrySet().iterator().next());

// Execute nested mail merge
document.getMailMerge().executeWidthNestedRegion("data_path", list);

// Save and dispose document
document.saveToFile("output_path", FileFormat.Docx);
document.dispose();
```

---

# Spire.Doc Bookmark Content Copy
## Copy content from a bookmark in a Word document
```java
// Get the bookmark named "Test" from the document
Bookmark bookmark = doc.getBookmarks().get("Test");

// Declare a DocumentObject variable to store the parent object of the bookmark start/end
DocumentObject docObj = null;

// Check if the paragraph containing the bookmark is within a cell of a table
// If it is within a cell, find the outermost parent object (Table) and get its start/end index on the document body
if (((Paragraph)bookmark.getBookmarkStart().getOwner()).isInCell()) {
    docObj = bookmark.getBookmarkStart().getOwner().getOwner().getOwner().getOwner();
} else {
    docObj = bookmark.getBookmarkStart().getOwner();
}

// Get the start index of the parent object on the document body
int startIndex = doc.getSections().get(0).getBody().getChildObjects().indexOf(docObj);

// Check if the paragraph containing the bookmark end is within a cell of a table
// If it is within a cell, find the outermost parent object (Table) and get its start/end index on the document body
if (((Paragraph)bookmark.getBookmarkEnd().getOwner()).isInCell()) {
    docObj = bookmark.getBookmarkEnd().getOwner().getOwner().getOwner().getOwner();
} else {
    docObj = bookmark.getBookmarkEnd().getOwner();
}

// Get the end index of the parent object on the document body
int endIndex = doc.getSections().get(0).getBody().getChildObjects().indexOf(docObj);

// Get the paragraph containing the bookmark start and its index within the paragraph
Paragraph para = (Paragraph)bookmark.getBookmarkStart().getOwner();
int pStartIndex = para.getChildObjects().indexOf(bookmark.getBookmarkStart());

// Get the paragraph containing the bookmark end and its index within the paragraph
para = (Paragraph)bookmark.getBookmarkEnd().getOwner();
int pEndIndex = para.getChildObjects().indexOf(bookmark.getBookmarkEnd());

// Create a TextBodySelection based on the start and end indices of the parent object and paragraph indices
TextBodySelection select =
    new TextBodySelection(doc.getSections().get(0).getBody(), startIndex, endIndex, pStartIndex, pEndIndex);

// Create a TextBodyPart using the TextBodySelection
TextBodyPart body = new TextBodyPart(select);

// Iterate through the items in the TextBodyPart and add them to the document's body
for (int i = 0; i < body.getBodyItems().getCount(); i++) {
    doc.getSections().get(0).getBody().getChildObjects().add(body.getBodyItems().get(i).deepClone());
}
```

---

# Spire.Doc Bookmark Creation
## Create simple and nested bookmarks in Word document
```java
// Create a simple bookmark
paragraph.appendBookmarkStart("SimpleCreateBookmark");
paragraph.appendText("This is a simple bookmark.");
paragraph.appendBookmarkEnd("SimpleCreateBookmark");

// Create nested bookmarks
paragraph.appendBookmarkStart("Root");
txtRange = paragraph.appendText(" This is Root data. ");
txtRange.getCharacterFormat().setItalic(true);

paragraph.appendBookmarkStart("NestedLevel1");
txtRange = paragraph.appendText(" This is Nested Level1. ");
txtRange.getCharacterFormat().setItalic(true);
txtRange.getCharacterFormat().setTextColor(new Color(40, 79, 79));

paragraph.appendBookmarkStart("NestedLevel2");
txtRange = paragraph.appendText(" This is Nested Level2. ");
txtRange.getCharacterFormat().setItalic(true);
txtRange.getCharacterFormat().setTextColor(new Color(105, 105, 105));

paragraph.appendBookmarkEnd("NestedLevel2");
paragraph.appendBookmarkEnd("NestedLevel1");
paragraph.appendBookmarkEnd("Root");
```

---

# Spire.Doc Table Bookmark Creation
## Create bookmark for a table in Word document
```java
// Define the createBookmarkForTable method
private static void createBookmarkForTable(Document doc, Section section) {
    // Add a paragraph to the section
    Paragraph paragraph = section.addParagraph();

    // Append text to the paragraph and set its formatting
    TextRange txtRange = paragraph
        .appendText("The following example demonstrates how to create bookmark for a table in a Word document.");
    txtRange.getCharacterFormat().setItalic(true);

    // Append bookmark start to the paragraph
    paragraph.appendBookmarkStart("CreateBookmark");

    // Append bookmark end to the paragraph
    paragraph.appendBookmarkEnd("CreateBookmark");

    // Add a table to the section
    Table table = section.addTable(true);

    // Reset the cells of the table to have 2 rows and 2 columns
    table.resetCells(2, 2);

    // Add text to the cells of the table
    TextRange range = table.getRows().get(0).getCells().get(0).addParagraph().appendText("sampleA");
    range = table.getRows().get(0).getCells().get(1).addParagraph().appendText("sampleB");
    range = table.getRows().get(1).getCells().get(0).addParagraph().appendText("120");
    range = table.getRows().get(1).getCells().get(1).addParagraph().appendText("260");

    // Get the first bookmark from the document
    Bookmark bookmark = doc.getBookmarks().get(0);

    // Get the name of the bookmark
    String bookmarkName = bookmark.getName();

    // Create a BookmarksNavigator for the document
    BookmarksNavigator navigator = new BookmarksNavigator(doc);

    // Move to the bookmark position in the document
    navigator.moveToBookmark(bookmarkName);

    // Get the TextBodyPart containing the bookmark content
    TextBodyPart part = navigator.getBookmarkContent();

    // Add the table to the BodyItems of the TextBodyPart
    part.getBodyItems().add(table);

    // Replace the bookmark content with the modified TextBodyPart
    navigator.replaceBookmarkContent(part);
}
```

---

# Spire.Doc Bookmark Text Extraction
## Extract text content from a specific bookmark in a Word document
```java
// Create a new instance of the Document class
Document doc = new Document();

// Create a BookmarksNavigator for the document
BookmarksNavigator navigator = new BookmarksNavigator(doc);

// Move to the bookmark position named "Test2" in the document
navigator.moveToBookmark("Test2");

// Get the TextBodyPart containing the bookmark content
TextBodyPart textBodyPart = navigator.getBookmarkContent();

// Iterate through the body items of the TextBodyPart
for (int i = 0; i < textBodyPart.getBodyItems().getCount(); i++) {
    // Check if the body item is a paragraph
    if (textBodyPart.getBodyItems().get(i) instanceof Paragraph) {
        // Cast the body item to a Paragraph
        Paragraph itemPara = (Paragraph)textBodyPart.getBodyItems().get(i);

        // Iterate through the child objects of the paragraph
        for (int j = 0; j < itemPara.getChildObjects().getCount(); j++) {
            // Check if the child object is a TextRange
            if (itemPara.getChildObjects().get(j) instanceof TextRange) {
                // Cast the child object to a TextRange
                TextRange textrange = (TextRange)(itemPara.getChildObjects().get(j));

                // Get the text from the TextRange
                String text = textrange.getText();
            }
        }
    }
}
```

---

# Get Bookmarks from Document
## Demonstrates how to get bookmarks from a Word document by index and by name
```java
// Create a new instance of the Document class
Document document = new Document();

// Get the first bookmark in the document by index
Bookmark bookmark1 = document.getBookmarks().get(0);

// Get the bookmark named "Test2" from the document
Bookmark bookmark2 = document.getBookmarks().get("Test2");

// Dispose of the document object to release resources
document.dispose();
```

---

# Spire.Doc Bookmark Extraction
## Get non-hidden bookmarks from a Word document
```java
// Retrieve all bookmark collections
BookmarkCollection collection = doc.getBookmarks();

// Initialize an empty string to store the extracted text
String text = "";

// Get bookmark name
if (collection != null && collection.getCount() > 0)
{
    for (Object object : collection)
    {
        Bookmark bookmark = (Bookmark) object;
        String name = bookmark.getName();
        if(!bookmark.isHidden())
        {
            text += String.format("The bookmark name is : " + name +"\n");
        }
        else
        {
            text += String.format("The hidden bookmark name is : " + name +"\n");
        }
    }
}
```

---

# Spire.Doc Insert Document at Bookmark
## Insert content from one document at a bookmark location in another document
```java
// Get the first section of the first document
Section section1 = document1.getSections().get(0);

// Create a BookmarksNavigator for the first document
BookmarksNavigator bn = new BookmarksNavigator(document1);

// Move to the bookmark position named "Test2" in the first document, preserving formatting and expanding the selection
bn.moveToBookmark("Test2", true, true);

// Get the BookmarkStart object at the current bookmark position
BookmarkStart start = bn.getCurrentBookmark().getBookmarkStart();

// Get the owner paragraph of the BookmarkStart
Paragraph para = start.getOwnerParagraph();

// Get the index of the owner paragraph within the body of the first document's section
int index = section1.getBody().getChildObjects().indexOf(para);

// Iterate through the sections and paragraphs of the second document
for (int i = 0; i < document2.getSections().getCount(); i++) {
    for (int j = 0; j < document2.getSections().get(i).getParagraphs().getCount(); j++) {
        // Deep clone each paragraph from the second document
        Paragraph insertPara = (Paragraph)document2.getSections().get(i).getParagraphs().get(j).deepClone();
        
        // Insert the cloned paragraph into the body of the first document's section, incrementing the index
        section1.getBody().getChildObjects().insert(index++ + 1, insertPara);
    }
}
```

---

# Spire.Doc Image Insertion at Bookmark
## Insert an image at a bookmark location in a Word document
```java
// Create a BookmarksNavigator object using the document
BookmarksNavigator bn = new BookmarksNavigator(doc);

// Move to the bookmark named "Test2"
bn.moveToBookmark("Test2", true, true);

// Create a new Section object
Section section0 = doc.addSection();

// Create a new Paragraph object
Paragraph paragraph = section0.addParagraph();

// Append the picture to the paragraph
DocPicture picture = paragraph.appendPicture("image_path");

// Insert the paragraph at the current bookmark position
bn.insertParagraph(paragraph);

// Remove the section from the document
doc.getSections().remove(section0);
```

---

# Spire.Doc Bookmark Removal
## Core functionality for removing a bookmark from a Word document
```java
// Get the bookmark named "Test2" from the document's bookmarks collection
Bookmark bookmark = document.getBookmarks().get("Test2");

// Remove the retrieved bookmark from the document's bookmarks collection
document.getBookmarks().remove(bookmark);
```

---

# Remove Bookmark Content
## This code demonstrates how to remove content between bookmark start and end positions in a Word document.
```java
// Get the bookmark named "Test2" from the document's bookmarks collection
Bookmark bookmark = document.getBookmarks().get("Test2");

// Get the Paragraph that contains the start position of the bookmark
Paragraph para = (Paragraph)bookmark.getBookmarkStart().getOwner();

// Determine the index of the bookmark start within its parent paragraph
int startIndex = para.getChildObjects().indexOf(bookmark.getBookmarkStart());

// Get the Paragraph that contains the end position of the bookmark
para = (Paragraph)bookmark.getBookmarkEnd().getOwner();

// Determine the index of the bookmark end within its parent paragraph
int endIndex = para.getChildObjects().indexOf(bookmark.getBookmarkEnd());

// Remove the content between the bookmark start and end positions
for (int i = startIndex + 1; i < endIndex; i++) {
    para.getChildObjects().removeAt(startIndex + 1);
}
```

---

# Spire.Doc Bookmark Replacement
## Replace content of a bookmark in a Word document
```java
// Create a BookmarksNavigator instance using the loaded document
BookmarksNavigator bookmarkNavigator = new BookmarksNavigator(doc);

// Move the navigator to the bookmark named "Test2"
bookmarkNavigator.moveToBookmark("Test2");

// Replace the content of the bookmark with the specified replacement content
bookmarkNavigator.replaceBookmarkContent("This is replaced content.", false);
```

---

# Replace Bookmark with Table
## Replace a bookmark in a Word document with a table using Spire.Doc for Java
```java
// Create a new Document instance
Document doc = new Document();

// Create a new table instance with auto-fit behavior
Table table = new Table(doc, true);

// Reset table cells with specified rows and columns
table.resetCells(3, 2);

// Populate the table with data
for (int i = 0; i < 3; i++) {
    for (int j = 0; j < 2; j++) {
        // Add a new paragraph to each cell and append the corresponding data
        table.getRows().get(i).getCells().get(j).addParagraph().appendText("Cell " + i + "," + j);
    }
}

// Create a BookmarksNavigator instance using the document
BookmarksNavigator navigator = new BookmarksNavigator(doc);

// Move the navigator to the bookmark named "Test2"
navigator.moveToBookmark("Test2");

// Create a TextBodyPart instance for the bookmark replacement content
TextBodyPart part = new TextBodyPart(doc);
part.getBodyItems().add(table);

// Replace the bookmark content with the new table
navigator.replaceBookmarkContent(part);
```

---

# Spire.Doc Bookmark Color Customization
## Set bookmark colors and styles in PDF conversion
```java
// Create a ToPdfParameterList instance to specify the PDF conversion settings
ToPdfParameterList toPdf = new ToPdfParameterList();

// Enable the creation of Word bookmarks in the PDF
toPdf.setCreateWordBookmarks(true);

// Set the title of the Word bookmarks in the PDF
toPdf.setWordBookmarksTitle("Changed bookmark");

// Set the color of the Word bookmarks to gray
toPdf.setWordBookmarksColor(Color.gray);

// Set the bookmark layout handler for the document
doc.BookmarkLayout = new com.spire.doc.documents.rendering.BookmarkLevelHandler() {
    @Override
    public void invoke(Object sender, com.spire.doc.documents.rendering.BookmarkLevelEventArgs args) {
        // Customize bookmark appearance based on its level
        if (args.getBookmarkLevel().getLevel() == 2) {
            args.getBookmarkLevel().setColor(Color.red);
            args.getBookmarkLevel().setStyle(BookmarkTextStyle.Bold);
        } else if (args.getBookmarkLevel().getLevel() == 3) {
            args.getBookmarkLevel().setColor(Color.gray);
            args.getBookmarkLevel().setStyle(BookmarkTextStyle.Italic);
        } else {
            args.getBookmarkLevel().setColor(Color.green);
            args.getBookmarkLevel().setStyle(BookmarkTextStyle.Regular);
        }
    }
};
```

---

# Spire.Doc Comment Functionality
## Add comment for specific text in a document
```java
// Custom method to insert comments into the document
private static void InsertComments(Document doc, String keystring) {
    // Find the text selection matching the keystring in the document
    TextSelection find = doc.findString(keystring, false, true);

    // Create a comment mark for the start of the comment
    CommentMark commentMarkStart = new CommentMark(doc, 1, CommentMarkType.Comment_Start);

    // Create a comment mark for the end of the comment
    CommentMark commentMarkEnd = new CommentMark(doc, 1, CommentMarkType.Comment_End);

    // Create a comment with the desired content and author
    Comment comment = new Comment(doc);
    comment.getBody().addParagraph().setText("Test comments");
    comment.getFormat().setAuthor("E-iceblue");

    // Get the text range as a single range
    TextRange range = find.getAsOneRange();

    // Get the owner paragraph of the text range
    Paragraph para = range.getOwnerParagraph();

    // Get the index of the text range in the child objects of the paragraph
    int index = para.getChildObjects().indexOf(range);

    // Add the comment as a child object of the paragraph
    para.getChildObjects().add(comment);

    // Insert the comment mark at the appropriate index in the paragraph's child objects
    para.getChildObjects().insert(index, commentMarkStart);
    para.getChildObjects().insert(index + 2, commentMarkEnd);
}
```

---

# Word Document Comment Handling
## Insert comments into a Word document using Spire.Doc
```java
// Create a new Document instance
Document document = new Document();

// Load the document from the specified input file
document.loadFromFile(input);

// Call the InsertComments method to insert comments in the first section of the document
InsertComments(document.getSections().get(0));

// Custom method to insert comments into a section
private static void InsertComments(Section section) {
    // Get the second paragraph in the section
    Paragraph paragraph = section.getParagraphs().get(1);

    // Append a comment to the paragraph with the specified text
    Comment comment = paragraph.appendComment("Spire.Doc for java");

    // Set the author of the comment
    comment.getFormat().setAuthor("E-iceblue");

    // Set the initial of the comment
    comment.getFormat().setInitial("CM");
}
```

---

# Spire.Doc Comment Extraction
## Extract text content from comments in a Word document
```java
// Create a new Document instance
Document doc = new Document();

// Iterate over each comment in the document
for (int i = 0; i < doc.getComments().getCount(); i++) {
    // Get the comment at the current index
    Comment comment = doc.getComments().get(i);

    // Iterate over each paragraph in the comment's body
    for (int j = 0; j < comment.getBody().getParagraphs().getCount(); j++) {
        // Get the paragraph at the current index
        Paragraph para = comment.getBody().getParagraphs().get(j);

        // Get the text of the paragraph
        String result = para.getText();
    }
}
```

---

# Spire.Doc Picture in Comment
## Insert a picture into a document comment
```java
// Get the third paragraph in the first section of the document
Paragraph paragraph = doc.getSections().get(0).getParagraphs().get(2);

// Append a comment to the paragraph with the specified text
Comment comment = paragraph.appendComment("This is a comment.");

// Set the author of the comment
comment.getFormat().setAuthor("E-iceblue");

// Create a new DocPicture instance for the document
DocPicture docPicture = new DocPicture(doc);

// Load the image from the specified input file into the DocPicture
docPicture.loadImage(input2);

// Add the DocPicture as a child object of a paragraph within the comment's body
comment.getBody().addParagraph().getChildObjects().add(docPicture);
```

---

# Spire.Doc Comment Management
## Demonstrates how to replace text in a comment and remove a comment from a document
```java
// Access the first comment, its body, and the first paragraph in the body
// Replace the text "This is the title" with "This comment is changed."
doc.getComments().get(0).getBody().getParagraphs().get(0).replace("This is the title", "This comment is changed.", false, false);

// Remove the comment at index 1 (the second comment)
doc.getComments().removeAt(1);
```

---

# Remove Content with Comment in Word Document
## Remove text content associated with a specific comment in a Word document
```java
// Get the second comment in the document
Comment comment = document.getComments().get(1);

// Get the paragraph that contains the obtained comment
Paragraph para = comment.getOwnerParagraph();

// Find the index of the CommentMarkStart in the paragraph's child objects
int startIndex = para.getChildObjects().indexOf(comment.getCommentMarkStart());

// Find the index of the CommentMarkEnd in the paragraph's child objects
int endIndex = para.getChildObjects().indexOf(comment.getCommentMarkEnd());

// Create an ArrayList to store TextRange objects
ArrayList<TextRange> list = new ArrayList<TextRange>();

// Iterate over the child objects between startIndex and endIndex
for (int i = startIndex; i < endIndex; i++) {
    // Check if the current child object is an instance of TextRange
    if (para.getChildObjects().get(i) instanceof TextRange) {
        // Add the TextRange object to the list
        list.add((TextRange) para.getChildObjects().get(i));
    }
}

// Create a new TextRange object associated with the document
TextRange textRange = new TextRange(document);

// Set the text of the new TextRange to null
textRange.setText(null);

// Insert the new TextRange object at the endIndex position in the paragraph's child objects
para.getChildObjects().insert(endIndex, textRange);

// Remove the previous TextRange objects from the paragraph's child objects
for (int i = 0; i < list.size(); i++) {
    para.getChildObjects().remove(list.get(i));
}
```

---

# Spire.Doc Java Comment Reply
## Reply to a comment in a Word document and add text and image to the reply
```java
// Get the first comment in the document
Comment comment1 = doc.getComments().get(0);

// Create a new Comment object associated with the document
Comment replyComment1 = new Comment(doc);

// Set the author of the reply comment
replyComment1.getFormat().setAuthor("E-iceblue");

// Add a paragraph to the body of the reply comment and append text to it
replyComment1.getBody().addParagraph().appendText("Spire.Doc for Java is a professional Word Java library on operating Word documents.");

// Make the reply comment a reply to comment1
comment1.replyToComment(replyComment1);

// Create a new DocPicture object associated with the document
DocPicture docPicture = new DocPicture(doc);

// Add the DocPicture object to the child objects of the first paragraph in the body of replyComment1
replyComment1.getBody().getParagraphs().get(0).getChildObjects().add(docPicture);
```

---

# Spire.Doc Barcode Image Addition
## Code to add a barcode image to a Word document
```java
// Add a paragraph to the first section of the document and append a picture to it
DocPicture picture = document.getSections().get(0).addParagraph().appendPicture("barcode.png");
```

---

# Spire.Doc Image and Shape
## Add image to document footer
```java
// Add a paragraph to the footer of the first section of the document and append a picture to it
DocPicture picture = document.getSections().get(0).getHeadersFooters().getFooter().addParagraph().appendPicture("image_path");

// Set the vertical origin of the picture to Page
picture.setVerticalOrigin(VerticalOrigin.Page);

// Set the horizontal origin of the picture to Page
picture.setHorizontalOrigin(HorizontalOrigin.Page);

// Set the vertical alignment of the picture to Bottom
picture.setVerticalAlignment(ShapeVerticalAlignment.Bottom);

// Set the text wrapping style of the picture to None
picture.setTextWrappingStyle(TextWrappingStyle.None);

// Add a TextBox to the footer of the first section of the document with specified width and height
TextBox textbox = document.getSections().get(0).getHeadersFooters().getFooter().addParagraph().appendTextBox(150, 20);

// Set the vertical origin of the TextBox to Page
textbox.setVerticalOrigin(VerticalOrigin.Page);

// Set the horizontal origin of the TextBox to Page
textbox.setHorizontalOrigin(HorizontalOrigin.Page);

// Set the horizontal position of the TextBox
textbox.setHorizontalPosition(300);

// Set the vertical position of the TextBox
textbox.setVerticalPosition(800);

// Add a paragraph to the TextBox body and append text to it
textbox.getBody().addParagraph().appendText("Welcome to E-iceblue");
```

---

# Spire.Doc Shape Group
## Add a group of shapes to a Word document
```java
// Create a new Document object
Document doc = new Document();

// Add a new Section to the document
Section sec = doc.addSection();

// Add a Paragraph to the Section
Paragraph para = sec.addParagraph();

// Append a ShapeGroup to the Paragraph with specified width and height
ShapeGroup shapegroup = para.appendShapeGroup(375, 462);
shapegroup.setHorizontalPosition(180);

// Calculate X scaling factor
float X = (float)(shapegroup.getWidth() / 1000.0f);

// Calculate Y scaling factor
float Y = (float)(shapegroup.getHeight() / 1000.0f);

// Create a TextBox with the document
TextBox txtBox = new TextBox(doc);

// Set the shape type of the TextBox to Round_Rectangle
txtBox.setShapeType(ShapeType.Round_Rectangle);

// Set the width of the TextBox
txtBox.setWidth(125 / X);

// Set the height of the TextBox
txtBox.setHeight(54 / Y);

// Add a paragraph to the TextBox body
Paragraph paragraph = txtBox.getBody().addParagraph();

// Set the horizontal alignment of the TextBox paragraph
paragraph.getFormat().setHorizontalAlignment(HorizontalAlignment.Center);

// Append text to the TextBox paragraph
paragraph.appendText("Start");

// Set the horizontal position of the TextBox
txtBox.setHorizontalPosition(19/ X);

// Set the vertical position of the TextBox
txtBox.setVerticalPosition(27 / Y);

// Set the line color of the TextBox
txtBox.getFormat().setLineColor(Color.GREEN);

// Add the TextBox to the child objects of the ShapeGroup
shapegroup.getChildObjects().add(txtBox);

// Create a ShapeObject with the document using Down_Arrow shape type
ShapeObject arrowLineShape = new ShapeObject(doc, ShapeType.Down_Arrow);

// Set the width of the arrow shape
arrowLineShape.setWidth(16 / X);

// Set the height of the arrow shape
arrowLineShape.setHeight(40 / Y);

// Set the horizontal position of the arrow shape
arrowLineShape.setHorizontalPosition(69 / X);

// Set the vertical position of the arrow shape
arrowLineShape.setVerticalPosition(87 / Y);

// Set the stroke color of the arrow shape
arrowLineShape.setStrokeColor(Color.PINK);

// Add the arrow shape to the child objects of the ShapeGroup
shapegroup.getChildObjects().add(arrowLineShape);

// Repeat the above steps with different values to create additional TextBoxes and shapes
txtBox = new TextBox(doc);
txtBox.setShapeType(ShapeType.Rectangle);
txtBox.setWidth(125 / X);
txtBox.setHeight(54 / Y);
paragraph = txtBox.getBody().addParagraph();
paragraph.getFormat().setHorizontalAlignment(HorizontalAlignment.Center);
paragraph.appendText("Step 1");
txtBox.setHorizontalPosition(19/ X);
txtBox.setVerticalPosition( 131/ Y);
txtBox.getFormat().setLineColor(Color.BLUE);
shapegroup.getChildObjects().add(txtBox);

arrowLineShape = new ShapeObject(doc, ShapeType.Down_Arrow);
arrowLineShape.setWidth(16 / X);
arrowLineShape.setHeight(40 / Y);
arrowLineShape.setHorizontalPosition(69 / X);
arrowLineShape.setVerticalPosition(192 / Y);
arrowLineShape.setStrokeColor(Color.PINK);
shapegroup.getChildObjects().add(arrowLineShape);

txtBox = new TextBox(doc);
txtBox.setShapeType(ShapeType.Parallelogram);
txtBox.setWidth(149 / X);
txtBox.setHeight(59/ Y);
paragraph = txtBox.getBody().addParagraph();
paragraph.getFormat().setHorizontalAlignment(HorizontalAlignment.Center);
paragraph.appendText("Step 2");
txtBox.setHorizontalPosition(7 / X);
txtBox.setVerticalPosition(236/ Y);
txtBox.getFormat().setLineColor(Color.MAGENTA);
shapegroup.getChildObjects().add(txtBox);

arrowLineShape = new ShapeObject(doc, ShapeType.Down_Arrow);
arrowLineShape.setWidth(16 / X);
arrowLineShape.setHeight(40/ Y);
arrowLineShape.setHorizontalPosition(66 / X);
arrowLineShape.setVerticalPosition(300 / Y);
arrowLineShape.setStrokeColor(Color.PINK);
shapegroup.getChildObjects().add(arrowLineShape);

txtBox = new TextBox(doc);
txtBox.setShapeType(ShapeType.Rectangle);
txtBox.setWidth( 125 / X);
txtBox.setHeight(54 / Y);
paragraph = txtBox.getBody().addParagraph();
paragraph.getFormat().setHorizontalAlignment(HorizontalAlignment.Center);
paragraph.appendText("Step 3");
txtBox.setHorizontalPosition(19 / X);
txtBox.setVerticalPosition(345 / Y);
txtBox.getFormat().setLineColor(Color.BLUE);
shapegroup.getChildObjects().add(txtBox);
```

---

# Spire.Doc Shape Addition
## Add various shapes to a Word document with positioning and page breaks
```java
// Create a new Document object
Document doc = new Document();

// Add a new Section to the document
Section sec = doc.addSection();

// Add a new Paragraph to the Section
Paragraph para = sec.addParagraph();

// Initialize variables for positioning shapes and counting lines
int x = 60, y = 40, lineCount = 0;

// Get a Map of ShapeType enum values and corresponding integers
Map<ShapeType, Integer> shapeTypes = getShapeTypes();

// Iterate from 1 to 19 (excluding 20)
for (int i = 1; i < 20; i++) {
    // Check if the line count is divisible by 8 to determine if a page break is needed
    if (lineCount > 0 && lineCount % 8 == 0) {
        para.appendBreak(BreakType.Page_Break);
        x = 60;
        y = 40;
        lineCount = 0;
    }

    // Append a shape to the paragraph with specified dimensions and shape type
    ShapeObject shape = para.appendShape(50, 50, getShapeType(shapeTypes, i));

    // Set the horizontal origin, position, vertical origin, and position of the shape
    shape.setHorizontalOrigin(HorizontalOrigin.Page);
    shape.setHorizontalPosition(x);
    shape.setVerticalOrigin(VerticalOrigin.Page);
    shape.setVerticalPosition(y + 50);

    // Update the x coordinate for the next shape
    x = x + (int) shape.getWidth() + 50;

    // Check if a new line needs to start based on the number of shapes added
    if (i > 0 && i % 5 == 0) {
        y = y + (int) shape.getHeight() + 120;
        lineCount++;
        x = 60;
    }
}

// Get the ShapeType enum value based on the given integer value from the Map
private static ShapeType getShapeType(Map<ShapeType, Integer> types, int value) {
    for (Map.Entry<ShapeType, Integer> entry : types.entrySet()) {
        if (entry.getValue().intValue() == value) {
            return entry.getKey();
        }
    }
    return null;
}

// Get a Map of ShapeType enum values and corresponding integers
private static Map<ShapeType, Integer> getShapeTypes() throws Exception {
    // Get all the ShapeType enum constants and fields
    Object[] enums = ShapeType.class.getEnumConstants();
    Field[] fields = ShapeType.class.getDeclaredFields();

    // Create a Map to store ShapeType and corresponding integer values
    Map<ShapeType, Integer> map = new HashMap<ShapeType, Integer>();

    // Iterate through the fields
    for (int i = 0; i < fields.length; i++) {
        // Skip final fields
        if (Modifier.isFinal(fields[i].getModifiers())) {
            continue;
        }

        // Set the field accessible
        fields[i].setAccessible(true);

        // Check if the field type is int
        if (fields[i].getType() == int.class) {
            // Loop through the enum constants
            for (int j = 0; j < enums.length; j++) {
                // Get the integer value of the field and put it in the map
                Object o = fields[i].get(enums[j]);
                map.put(((ShapeType) enums[j]), (Integer) o);
            }
        }
    }
    return map;
}
```

---

# Spire.Doc SVG Image Addition
## Add SVG image to Word document
```java
// Create a new Document object
Document document = new Document();

// Add a new Section to the document
Section section = document.addSection();

// Add a new Paragraph to the section
Paragraph paragraph = section.addParagraph();

// Append the picture (SVG) to the paragraph
paragraph.appendPicture(/* SVG file path */);
```

---

# Spire.Doc Shape Alignment
## Align shapes to center in a Word document
```java
// Create a new Document object
Document doc = new Document();

// Get the first section of the document
Section section = doc.getSections().get(0);

// Iterate through each paragraph in the section
for (int i=0; i<section.getParagraphs().getCount(); i++) {
    // Get the current paragraph
    Paragraph para = section.getParagraphs().get(i);

    // Iterate through each child object in the paragraph
    for (int j=0; j<para.getChildObjects().getCount(); j++) {
        // Get the current document object
        DocumentObject obj = para.getChildObjects().get(j);

        // Check if the object is an instance of ShapeObject
        if (obj instanceof ShapeObject) {
            // Set the horizontal alignment of the shape to Center
            ((ShapeObject) obj).setHorizontalAlignment(ShapeHorizontalAlignment.Center);
        }
    }
}
```

---

# Spire.Doc Image Extraction
## Extract images from a Word document using breadth-first search traversal
```java
// Create a new Document object and load it from a file
Document document = new Document();

// Create a Queue to store composite objects
Queue<ICompositeObject> nodes = new LinkedList<ICompositeObject>();

// Add the document as the first node in the queue
nodes.add(document);

// Create a List to store extracted images
List<BufferedImage> images = new ArrayList<BufferedImage>();

// Traverse the document tree using breadth-first search
while (nodes.size() > 0) {
    // Retrieve and remove the next node from the queue
    ICompositeObject node = nodes.poll();

    // Iterate through the child objects of the current node
    for (int i = 0; i < node.getChildObjects().getCount(); i++) {
        // Get the current child object
        IDocumentObject child = node.getChildObjects().get(i);

        // Check if the child object is a composite object
        if (child instanceof ICompositeObject) {
            // Add the composite object to the queue for further traversal
            nodes.add((ICompositeObject) child);
            if (child.getDocumentObjectType() == DocumentObjectType.Picture) 
         {
                // Extract the image from the picture and add it to the list
                DocPicture picture = (DocPicture) child;
                images.add(picture.getImage());
            }
        }
    }
}
```

---

# Spire.Doc Image Handling
## Insert image into Word document
```java
// Add a new Paragraph to the section and set horizontal alignment to Left
Paragraph paragraph = section.addParagraph();
paragraph.getFormat().setHorizontalAlignment(HorizontalAlignment.Left);

// Append the picture to the paragraph with specified width and height
DocPicture picture = paragraph.appendPicture(input);
picture.setWidth(100f);
picture.setHeight(100f);
```

---

# Spire.Doc Insert Image
## Insert an image into a Word document with specific formatting
```java
// Get the first section of the document
Section section = doc.getSections().get(0);

// Get the first paragraph of the section if it exists, otherwise add a new paragraph
Paragraph paragraph = section.getParagraphs().getCount() > 0 ? section.getParagraphs().get(0) : section.addParagraph();

// Append text to the paragraph
paragraph.appendText("The sample demonstrates how to insert an image into a document.");

// Apply a built-in style (Heading 2) to the paragraph
paragraph.applyStyle(BuiltinStyle.Heading_2);

// Add a new paragraph to the section
paragraph = section.addParagraph();

// Append text to the new paragraph
paragraph.appendText("This is an inserted picture.");

// Create a new DocPicture object and load the image
DocPicture picture = new DocPicture(doc);
picture.loadImage(input2);

// Set the horizontal and vertical position of the picture
picture.setHorizontalPosition(50.0F);
picture.setVerticalPosition(60.0F);

// Set the width and height of the picture
picture.setWidth(200);
picture.setHeight(200);

// Set the text wrapping style of the picture
picture.setTextWrappingStyle(TextWrappingStyle.Through);

// Insert the picture at the beginning of the paragraph's child objects
paragraph.getChildObjects().insert(0, picture);
```

---

# spire.doc word art insertion
## Insert WordArt into a document
```java
// Create a new Document object
Document doc = new Document();

// Get the first section of the document and add a new paragraph to it
Paragraph paragraph = doc.getSections().get(0).addParagraph();

// Append a shape to the paragraph with specified dimensions and shape type (Text_Wave_4)
ShapeObject shape = paragraph.appendShape(250, 70, ShapeType.Text_Wave_4);

// Set the vertical and horizontal position of the shape
shape.setVerticalPosition(20);
shape.setHorizontalPosition(80);

// Set the text content of the WordArt in the shape
shape.getWordArt().setText("Spire.Doc for JAVA");

// Set the fill color of the shape to red
shape.setFillColor(Color.RED);

// Set the stroke color of the shape to yellow
shape.setStrokeColor(Color.YELLOW);
```

---

# Remove Shapes from Word Document
## This code demonstrates how to remove shape objects from a Word document by traversing through paragraphs and their child objects.
```java
// Get the first section of the document
Section section = doc.getSections().get(0);

// Traverse all child objects of each paragraph in the section
for (int j = 0; j < section.getParagraphs().getCount(); j++) {
    Paragraph para = section.getParagraphs().get(j);
    for (int i = 0; i < para.getChildObjects().getCount(); i++) {
        // Check if the child object is a ShapeObject
        if (para.getChildObjects().get(i) instanceof ShapeObject) {
            // Remove the ShapeObject from the paragraph's child objects
            para.getChildObjects().removeAt(i);
        }
    }
}
```

---

# Spire.Doc Image Replacement
## Replace images with text in a Word document
```java
// Initialize a counter for image labels
int j = 1;

// Iterate through all sections of the document
for (int i = 0; i < doc.getSections().getCount(); i++) {
    Section sec = doc.getSections().get(i);
    
    // Iterate through all paragraphs in the section
    for (int p = 0; p < sec.getParagraphs().getCount(); p++) {
        Paragraph para = sec.getParagraphs().get(p);
        
        // Create an ArrayList to store picture objects found in the paragraph
        ArrayList<DocumentObject> pictures = new ArrayList<DocumentObject>();
        
        // Iterate through all child objects of the paragraph
        for (int o = 0; o < para.getChildObjects().getCount(); o++) {
            DocumentObject docObj = para.getChildObjects().get(o);
            
            // Check if the child object is a Picture
            if (docObj.getDocumentObjectType() == DocumentObjectType.Picture) {
                pictures.add(docObj);
            }
        }
        
        // Iterate through all picture objects in the ArrayList
        for (int m = 0; m < pictures.size(); m++) {
            DocumentObject pic = pictures.get(m);
            int index = para.getChildObjects().indexOf(pic);
            
            // Create a new TextRange object with the image label and insert it before the picture object
            TextRange range = new TextRange(doc);
            range.setText(String.format("Here was image-%d", j));
            para.getChildObjects().insert(index, range);
            
            // Remove the picture object from the paragraph's child objects
            para.getChildObjects().remove(pic);
            
            // Increment the image label counter
            j++;
        }
    }
}
```

---

# Spire.Doc Image Resizing
## Reset image size in a Word document
```java
// Get the first section of the document
Section section = doc.getSections().get(0);

// Get the first paragraph of the section
Paragraph paragraph = section.getParagraphs().get(0);

// Iterate through all child objects of the paragraph
for (int i = 0; i < paragraph.getChildObjects().getCount(); i++) {
    DocumentObject docObj = paragraph.getChildObjects().get(i);
    
    // Check if the child object is an instance of DocPicture
    if (docObj instanceof DocPicture) {
        // Cast the child object to DocPicture and modify its width and height
        DocPicture picture = (DocPicture) docObj;
        picture.setWidth(50f);
        picture.setHeight(50f);
    }
}
```

---

# Spire.Doc Shape Size Reset
## Reset the size of a shape in a Word document
```java
// Get the first section of the document
Section section = doc.getSections().get(0);

// Get the first paragraph of the section
Paragraph para = section.getParagraphs().get(0);

// Get the second child object of the paragraph and cast it to ShapeObject
ShapeObject shape = (ShapeObject) para.getChildObjects().get(1);

// Set the width and height of the shape
shape.setWidth(200);
shape.setHeight(200);
```

---

# Spire.Doc Shape Rotation
## Rotate shapes in a Word document
```java
// Get the first section of the document
Section section = doc.getSections().get(0);

// Iterate through each paragraph in the section
for (int i = 0; i < section.getParagraphs().getCount(); i++) {
    // Get the current paragraph
    Paragraph para = section.getParagraphs().get(i);
  
    // Iterate through each child object in the paragraph
    for (int j = 0; j < para.getChildObjects().getCount(); j++) {
        // Get the current child object
        DocumentObject obj = para.getChildObjects().get(j);
        
        // Check if the child object is a ShapeObject
        if (obj instanceof ShapeObject) {
            // Set the rotation of the ShapeObject to 20.0
            ((ShapeObject) obj).setRotation(20.0);
        }
    }
}
```

---

# Spire.Doc Shape Styling
## Set the style properties of a line shape in a Word document
```java
// Append a shape (line) to the paragraph with specified dimensions and type
ShapeObject shape = paragraph.appendShape(100, 100, ShapeType.Action_Button_Blank);

// Set the fill color of the shape to orange
shape.getFill().setColor(Color.orange);

// Set the stroke color of the shape to black
shape.setStrokeColor(Color.black);

// Set the line style of the shape to single
shape.setLineStyle(ShapeLineStyle.Single);

// Set the line dashing style of the shape to long dash dot dot GEL
shape.setLineDashing(LineDashing.Long_Dash_Dot_Dot_GEL);
```

---

# Spire.Doc Text Wrapping
## Set text wrapping style for images in Word document
```java
// Iterate through each section in the document
for (int i = 0; i < doc.getSections().getCount(); i++) {
    // Get the current section
    Section sec = doc.getSections().get(i);

    // Iterate through each paragraph in the section
    for (int j = 0; j < sec.getParagraphs().getCount(); j++) {
        // Get the current paragraph
        Paragraph para = sec.getParagraphs().get(j);

        // Iterate through each child object in the paragraph
        for (int p = 0; p < para.getChildObjects().getCount(); p++) {
            // Get the current child object
            DocumentObject docObj = para.getChildObjects().get(p);

            // Check if the child object is a Picture
            if (docObj.getDocumentObjectType() == DocumentObjectType.Picture) {
                // Cast the child object to DocPicture
                DocPicture picture = (DocPicture) docObj;

                // Set the text wrapping style of the picture to Through
                picture.setTextWrappingStyle(TextWrappingStyle.Through);

                // Set the text wrapping type of the picture to Both
                picture.setTextWrappingType(TextWrappingType.Both);
            }
        }
    }
}
```

---

# Spire.Doc Image Transparency
## Set transparent color for images in a document
```java
// Get the first paragraph in the first section of the document
Paragraph paragraph = doc.getSections().get(0).getParagraphs().get(0);

// Iterate through each child object in the paragraph
for (int i = 0; i < paragraph.getChildObjects().getCount(); i++) {
    // Get the current child object
    DocumentObject obj = paragraph.getChildObjects().get(i);

    // Check if the child object is a DocPicture
    if (obj instanceof DocPicture) {
        // Cast the child object to DocPicture
        DocPicture picture = (DocPicture) obj;

        // Set the transparent color of the picture to blue
        picture.setTransparentColor(Color.BLUE);
    }
}
```

---

# Update Image in Word Document
## Find and replace an image in a Word document using Spire.Doc for Java
```java
// Create an ArrayList to store DocumentObject instances of pictures
ArrayList<DocumentObject> pictures = new ArrayList<>();

// Iterate through each section in the document
for (int i = 0; i < doc.getSections().getCount(); i++) {
    // Get the current section
    Section sec = doc.getSections().get(i);

    // Iterate through each paragraph in the section
    for (int p = 0; p < sec.getParagraphs().getCount(); p++) {
        // Get the current paragraph
        Paragraph para = sec.getParagraphs().get(p);

        // Iterate through each child object in the paragraph
        for (int o = 0; o < para.getChildObjects().getCount(); o++) {
            // Get the current child object
            DocumentObject docObj = para.getChildObjects().get(o);

            // Check if the child object is a Picture
            if (docObj.getDocumentObjectType() == DocumentObjectType.Picture) {
                // Add the picture object to the pictures ArrayList
                pictures.add(docObj);
            }
        }
    }
}

// Get the first picture from the pictures ArrayList
DocPicture picture = (DocPicture) pictures.get(0);

// Update the picture with a new image
picture.loadImage(input2);
```

---

# Spire.Doc First Page Header
## Add header only to the first page of a document
```java
// Get the header of the first section in doc1
HeaderFooter header = doc1.getSections().get(0).getHeadersFooters().getHeader();

// Get the first page header of the first section in doc2
HeaderFooter firstPageHeader = doc2.getSections().get(0).getHeadersFooters().getFirstPageHeader();

// Iterate through each section in doc2
for (int i = 0; i < doc2.getSections().getCount(); i++) {
    // Set different first page header/footer for each section in doc2
    doc2.getSections().get(i).getPageSetup().setDifferentFirstPageHeaderFooter(true);
}

// Clear the paragraphs in the firstPageHeader
firstPageHeader.getParagraphs().clear();

// Copy the child objects from header to firstPageHeader
for (int j = 0; j < header.getChildObjects().getCount(); j++) {
    // Get the current child object in header
    DocumentObject obj = header.getChildObjects().get(j);

    // Deep clone the child object and add it to firstPageHeader
    firstPageHeader.getChildObjects().add(obj.deepClone());
}
```

---

# Adjust Header and Footer Height
## This code demonstrates how to adjust the header and footer height in a Word document using Spire.Doc for Java.
```java
// Create a new Document object
Document doc = new Document();

// Get the first section of the document
Section section = doc.getSections().get(0);

// Set the header distance of the section to 100 points
section.getPageSetup().setHeaderDistance(100);

// Set the footer distance of the section to 100 points
section.getPageSetup().setFooterDistance(100);
```

---

# Spire.Doc Header Footer Copy
## Copy header from one document to multiple sections of another document
```java
// Get the header of the first section in source document
HeaderFooter header = doc1.getSections().get(0).getHeadersFooters().getHeader();

// Iterate through each section in target document
for (int i = 0; i < doc2.getSections().getCount(); i++) {
    // Iterate through each child object in the header
    for (int j = 0; j < header.getChildObjects().getCount(); j++) {
        // Get the current child object in the header
        DocumentObject obj = header.getChildObjects().get(j);

        // Deep clone the child object and add it to the header of the current section in target document
        doc2.getSections().get(i).getHeadersFooters().getHeader().getChildObjects().add(obj.deepClone());
    }
}
```

---

# Spire.Doc Different First Page Header/Footer
## Set up different headers and footers for the first page of a document
```java
// Get the first section and enable different first page header/footer
Section section = doc.getSections().get(0);
section.getPageSetup().setDifferentFirstPageHeaderFooter(true);

// Set the first page header and append a picture to it
Paragraph paragraph1 = section.getHeadersFooters().getFirstPageHeader().addParagraph();
paragraph1.getFormat().setHorizontalAlignment(HorizontalAlignment.Right);
paragraph1.appendPicture(input2);

// Set the first page footer
Paragraph paragraph2 = section.getHeadersFooters().getFirstPageFooter().addParagraph();
paragraph2.getFormat().setHorizontalAlignment(HorizontalAlignment.Center);
TextRange firstPageFooter = paragraph2.appendText("First Page Footer");
firstPageFooter.getCharacterFormat().setFontSize(10);

// Set the other header
Paragraph paragraph3 = section.getHeadersFooters().getHeader().addParagraph();
paragraph3.getFormat().setHorizontalAlignment(HorizontalAlignment.Center);
TextRange headerText = paragraph3.appendText("Spire.Doc for JAVA");
headerText.getCharacterFormat().setFontSize(10);

// Set the other footer
Paragraph paragraph4 = section.getHeadersFooters().getFooter().addParagraph();
paragraph4.getFormat().setHorizontalAlignment(HorizontalAlignment.Center);
TextRange footerText = paragraph4.appendText("E-iceblue");
footerText.getCharacterFormat().setFontSize(10);
```

---

# Spire.Doc Header and Footer
## Create and configure headers and footers in a Word document
```java
// Get the first section of the document
Section section = document.getSections().get(0);

// Set page size and margins for the section
section.getPageSetup().setPageSize(PageSize.A4);
section.getPageSetup().getMargins().setTop(72f);
section.getPageSetup().getMargins().setBottom(72f);
section.getPageSetup().getMargins().setLeft(89.85f);
section.getPageSetup().getMargins().setRight(89.85f);

// Get header and footer objects from the section
HeaderFooter header = section.getHeadersFooters().getHeader();
HeaderFooter footer = section.getHeadersFooters().getFooter();

// Insert picture and text to the header
Paragraph headerParagraph = header.addParagraph();
// Add image to header
DocPicture headerPicture = headerParagraph.appendPicture(/* header image file path */);

// Add header text
TextRange text = headerParagraph.appendText("Demo of Spire.Doc");
text.getCharacterFormat().setFontName("Arial");
text.getCharacterFormat().setFontSize(10);
text.getCharacterFormat().setItalic(true);
headerParagraph.getFormat().setHorizontalAlignment(HorizontalAlignment.Right);

// Set border properties for the header paragraph
headerParagraph.getFormat().getBorders().getBottom().setBorderType(BorderStyle.Single);
headerParagraph.getFormat().getBorders().getBottom().setSpace(0.05F);

// Set layout properties for the header picture
headerPicture.setTextWrappingStyle(TextWrappingStyle.Behind);
headerPicture.setHorizontalOrigin(HorizontalOrigin.Page);
headerPicture.setHorizontalAlignment(ShapeHorizontalAlignment.Left);
headerPicture.setVerticalOrigin(VerticalOrigin.Page);
headerPicture.setVerticalAlignment(ShapeVerticalAlignment.Top);

// Insert picture to the footer
Paragraph footerParagraph = footer.addParagraph();
// Add image to footer
DocPicture footerPicture = footerParagraph.appendPicture(/* footer image file path */);

// Set layout properties for the footer picture
footerPicture.setTextWrappingStyle(TextWrappingStyle.Behind);
footerPicture.setHorizontalOrigin(HorizontalOrigin.Page);
footerPicture.setHorizontalAlignment(ShapeHorizontalAlignment.Left);
footerPicture.setVerticalOrigin(VerticalOrigin.Page);
footerPicture.setVerticalAlignment(ShapeVerticalAlignment.Bottom);

// Insert page number and total number of pages in the footer
footerParagraph.appendField("page number", FieldType.Field_Page);
footerParagraph.appendText(" of ");
footerParagraph.appendField("number of pages", FieldType.Field_Num_Pages);
footerParagraph.getFormat().setHorizontalAlignment(HorizontalAlignment.Right);

// Set border properties for the footer paragraph
footerParagraph.getFormat().getBorders().getTop().setBorderType(BorderStyle.Single);
footerParagraph.getFormat().getBorders().getTop().setSpace(0.05F);
```

---

# Spire.Doc Image Header and Footer
## Add images to document header and footer
```java
// Get the header of the first page
HeaderFooter header = doc.getSections().get(0).getHeadersFooters().getHeader();

// Add a paragraph to the header
Paragraph paragraph = header.addParagraph();

// Set the horizontal alignment of the paragraph
paragraph.getFormat().setHorizontalAlignment(HorizontalAlignment.Right);

// Append a picture to the paragraph
DocPicture headerImage = paragraph.appendPicture("image-path-for-header.jpg");
headerImage.setVerticalAlignment(ShapeVerticalAlignment.Bottom);

// Get the footer of the first section
HeaderFooter footer = doc.getSections().get(0).getHeadersFooters().getFooter();

// Add a paragraph to the footer
Paragraph paragraph2 = footer.addParagraph();

// Set the horizontal alignment of the paragraph
paragraph2.getFormat().setHorizontalAlignment(HorizontalAlignment.Left);

// Append a picture to the paragraph
DocPicture footerImage = paragraph2.appendPicture("image-path-for-footer.png");

// Append text to the paragraph
TextRange copyrightText = paragraph2.appendText("Copyright © 2013 e-iceblue. All Rights Reserved.");
copyrightText.getCharacterFormat().setFontName("Arial");
copyrightText.getCharacterFormat().setFontSize(10);
copyrightText.getCharacterFormat().setTextColor(Color.BLACK);
```

---

# Spire.Doc Header Locking
## Lock and unlock headers in a Word document
```java
// Get the first section of the document
Section section = doc.getSections().get(0);

// Protect the document and set the ProtectionType as AllowOnlyFormFields with password "123"
doc.protect(ProtectionType.Allow_Only_Form_Fields, "123");

// Set the ProtectForm property of the section as false to unprotect it
section.setProtectForm(false);
```

---

# Spire.Doc Header Footer Management
## Create different headers and footers for odd and even pages
```java
// Get the first section of the document
Section section = doc.getSections().get(0);

// Set different odd and even pages header/footer
section.getPageSetup().setDifferentOddAndEvenPagesHeaderFooter(true);

// Add a paragraph to the odd header
Paragraph P3 = section.getHeadersFooters().getOddHeader().addParagraph();
TextRange OH = P3.appendText("Odd Header");
P3.getFormat().setHorizontalAlignment(HorizontalAlignment.Center);
OH.getCharacterFormat().setFontName("Arial");
OH.getCharacterFormat().setFontSize(10);

// Add a paragraph to the even header
Paragraph P4 = section.getHeadersFooters().getEvenHeader().addParagraph();
TextRange EH = P4.appendText("Even Header from E-iceblue Using Spire.Doc");
P4.getFormat().setHorizontalAlignment(HorizontalAlignment.Center);
EH.getCharacterFormat().setFontName("Arial");
EH.getCharacterFormat().setFontSize(10);

// Add a paragraph to the odd footer
Paragraph P2 = section.getHeadersFooters().getOddFooter().addParagraph();
TextRange OF = P2.appendText("Odd Footer");
P2.getFormat().setHorizontalAlignment(HorizontalAlignment.Center);
OF.getCharacterFormat().setFontName("Arial");
OF.getCharacterFormat().setFontSize(10);

// Add a paragraph to the even footer
Paragraph P1 = section.getHeadersFooters().getEvenFooter().addParagraph();
TextRange EF = P1.appendText("Even Footer from E-iceblue Using Spire.Doc");
EF.getCharacterFormat().setFontName("Arial");
EF.getCharacterFormat().setFontSize(10);
P1.getFormat().setHorizontalAlignment(HorizontalAlignment.Center);
```

---

# Spire.Doc Page Border Configuration
## Configure page borders and their relationship with headers and footers
```java
// Create a new document object
Document doc = new Document();

// Add a section to the document
Section section = doc.addSection();

// Set page borders properties
section.getPageSetup().getBorders().setBorderType(BorderStyle.Wave);
section.getPageSetup().getBorders().setColor(Color.GREEN);
section.getPageSetup().getBorders().getLeft().setSpace(20);
section.getPageSetup().getBorders().getRight().setSpace(20);

// Add a paragraph to the header
Paragraph paragraph1 = section.getHeadersFooters().getHeader().addParagraph();
paragraph1.getFormat().setHorizontalAlignment(HorizontalAlignment.Right);
TextRange headerText = paragraph1.appendText("Header isn't included in page border");
headerText.getCharacterFormat().setFontName("Calibri");
headerText.getCharacterFormat().setFontSize(20);
headerText.getCharacterFormat().setBold(true);

// Add a paragraph to the footer
Paragraph paragraph2 = section.getHeadersFooters().getFooter().addParagraph();
paragraph2.getFormat().setHorizontalAlignment(HorizontalAlignment.Left);
TextRange footerText = paragraph2.appendText("Footer is included in page border");
footerText.getCharacterFormat().setFontName("Calibri");
footerText.getCharacterFormat().setFontSize(20);
footerText.getCharacterFormat().setBold(true);

// Set page setup properties
section.getPageSetup().setPageBorderIncludeHeader(false);
section.getPageSetup().setHeaderDistance(40);
section.getPageSetup().setPageBorderIncludeFooter(true);
section.getPageSetup().setFooterDistance(40);
```

---

# Spire.Doc Footer Removal
## Remove footers from a Word document
```java
// Get the first section of the document
Section section = doc.getSections().get(0);

HeaderFooter footer;

// Clear the child objects of the footer for the first page
footer = section.getHeadersFooters().getByHeaderFooterType(HeaderFooterType.Footer_First_Page);
if (footer != null) {
    footer.getChildObjects().clear();
}

// Clear the child objects of the footer for odd pages
footer = section.getHeadersFooters().getByHeaderFooterType(HeaderFooterType.Footer_Odd);
if (footer != null) {
    footer.getChildObjects().clear();
}

// Clear the child objects of the footer for even pages
footer = section.getHeadersFooters().getByHeaderFooterType(HeaderFooterType.Footer_Even);
if (footer != null) {
    footer.getChildObjects().clear();
}
```

---

# Remove Document Headers
## This code demonstrates how to remove headers from different types of pages in a Word document
```java
// Get the first section of the document
Section section = doc.getSections().get(0);

HeaderFooter header;

// Clear the child objects of the header for the first page
header = section.getHeadersFooters().getByHeaderFooterType(HeaderFooterType.Header_First_Page);
if (header != null) {
    header.getChildObjects().clear();
}

// Clear the child objects of the header for odd pages
header = section.getHeadersFooters().getByHeaderFooterType(HeaderFooterType.Header_Odd);
if (header != null) {
    header.getChildObjects().clear();
}

// Clear the child objects of the header for even pages
header = section.getHeadersFooters().getByHeaderFooterType(HeaderFooterType.Header_Even);
if (header != null) {
    header.getChildObjects().clear();
}
```

---

# Spire.Doc Table Alternative Text
## Add title and description to a table in Word document
```java
// Get the first section of the document
Section section = doc.getSections().get(0);

// Get the first table in the section
Table table = (Table) section.getTables().get(0);

// Set the title of the table
table.setTitle("Table 1");

// Set the description of the table
table.setTableDescription("Description Text");
```

---

# Spire.Doc Table Row Operations
## Demonstrates how to add and delete rows in a Word document table
```java
// Get the first section of the document
Section section = document.getSections().get(0);

// Get the first table in the section
Table table = section.getTables().get(0);

// Remove the row at index 7 from the table
table.getRows().removeAt(7);

// Create a new table row
TableRow row = new TableRow(document);

// Add cells to the new row with centered paragraphs containing "Added" text
for (int i = 0; i < table.getRows().get(0).getCells().getCount(); i++) {
    TableCell tc = row.addCell();
    Paragraph paragraph = tc.addParagraph();
    paragraph.getFormat().setHorizontalAlignment(HorizontalAlignment.Center);
    paragraph.appendText("Added");
}

// Insert the new row at index 2 in the table
table.getRows().insert(2, row);

// Add a new row at the end of the table
table.addRow();
```

---

# Spire.Doc Table Column Operations
## Add or remove columns from a table in a Word document
```java
// Method to add a column to the table at a specific columnIndex
private static void addColumn(Table table, int columnIndex) {
    for (int r = 0; r < table.getRows().getCount(); r++) {
        TableCell addCell = new TableCell(table.getDocument());
        table.getRows().get(r).getCells().insert(columnIndex, addCell);
    }
}

// Method to remove a column from the table at a specific columnIndex
private static void removeColumn(Table table, int columnIndex) {
    for (int r = 0; r < table.getRows().getCount(); r++) {
        table.getRows().get(r).getCells().removeAt(columnIndex);
    }
}
```

---

# spire.doc table cell picture
## add picture to a table cell in a Word document
```java
// Get the first table in the first section of the document
Table table1 = doc.getSections().get(0).getTables().get(0);

// Get the paragraph at row 1, cell 2 of table1 and append a picture
DocPicture picture = table1.getRows().get(1).getCells().get(2).getParagraphs().get(0).appendPicture(input2);

// Set the width and height of the picture to 100 units
picture.setWidth(100);
picture.setHeight(100);
```

---

# Spire.Doc Table Creation
## Create a table in a document using array data
```java
// Add a table to the section with headers enabled
Table table = section.addTable(true);

// Define the data for the table
String[][] data = {
        {"Name", "Capital", "Continent", "Area", "Population"},
        {"Argentina", "Buenos Aires", "South America", "2777815", "32300003"},
        {"Bolivia", "La Paz", "South America", "1098575", "7300000"},
        {"Brazil", "Brasilia", "South America", "8511196", "150400000"},
};

int rowCount = data.length;
int columnCount = data[0].length;

// Reset the cells of the table with the specified row and column count
table.resetCells(rowCount, columnCount);

// Populate the table with the data
for (int i = 0; i < rowCount; i++) {
    for (int j = 0; j < columnCount; j++) {
        table.getRows().get(i).getCells().get(j).addParagraph().appendText(data[i][j]);
    }
}
```

---

# Spire.Doc Table Row Formatting
## Allow table rows to break across pages in a Word document
```java
// Get the first section of the document
Section section = document.getSections().get(0);

// Get the first table in the section
Table table = section.getTables().get(0);

// Set the "break across pages" property to true for each row in the table
for (int i = 0; i < table.getRows().getCount(); i++) {
    TableRow row = table.getRows().get(i);
    row.getRowFormat().isBreakAcrossPages(true);
}
```

---

# Spire.Doc Table Auto-Fit
## Demonstrates how to auto-fit a table to its contents in a Word document
```java
// Get the first section of the document
Section section = document.getSections().get(0);

// Get the first table in the section
Table table = section.getTables().get(0);

// Auto-fit the table to its contents
table.autoFit(AutoFitBehaviorType.Auto_Fit_To_Contents);
```

---

# Table AutoFit with Fixed Column Widths
## Demonstrates how to auto-fit a table using fixed column widths in a Word document
```java
// Get the first section of the document
Section section = document.getSections().get(0);

// Get the first table in the section
Table table = section.getTables().get(0);

// Auto-fit the table using fixed column widths
table.autoFit(AutoFitBehaviorType.Fixed_Column_Widths);
```

---

# Spire.Doc Table Auto-Fit
## Auto-fit table to window in Word document
```java
// Get the first section of the document
Section section = document.getSections().get(0);

// Get the first table in the section
Table table = section.getTables().get(0);

// Auto-fit the table to fit the window
table.autoFit(AutoFitBehaviorType.Auto_Fit_To_Window);
```

---

# Spire.Doc Table Cloning
## Clone a table in a Word document
```java
// Get the first section of the document
Section se = doc.getSections().get(0);

// Get the original table from the section
Table original_Table = se.getTables().get(0);

// Deep clone the original table to create a copy
Table copied_Table = original_Table.deepClone();

// Add the copied table to the section
se.getTables().add(copied_Table);
```

---

# spire.doc table operations
## combine and split tables in a document
```java
private static void CombineTables() {
    // Get the first section
    Section section = doc.getSections().get(0);

    // Get the first and second table
    Table table1 = section.getTables().get(0);
    Table table2 = section.getTables().get(1);

    // Add the rows of table2 to table1
    for (int i = 0; i < table2.getRows().getCount(); i++) {
        table1.getRows().add(table2.getRows().get(i).deepClone());
    }

    // Remove table2
    section.getTables().remove(table2);
}

private static void SplitTable() {
    // Get the first section
    Section section = doc.getSections().get(0);

    // Get the first table
    Table table = section.getTables().get(0);

    // We will split the table at the third row
    int splitIndex = 2;

    // Create a new table for the split table
    Table newTable = new Table(section.getDocument());

    // Add rows to the new table
    for (int i = splitIndex; i < table.getRows().getCount(); i++) {
        newTable.getRows().add(table.getRows().get(i).deepClone());
    }

    // Remove rows from the original table
    for (int i = table.getRows().getCount() - 1; i >= splitIndex; i--) {
        table.getRows().removeAt(i);
    }

    // Add the new table in section
    section.getTables().add(newTable);
}
```

---

# Spire.Doc Nested Table Creation
## Create a nested table structure in a Word document
```java
// Create a new document
Document doc = new Document();
Section section = doc.addSection();

// Add a table
Table table = section.addTable(true);
table.resetCells(2, 2);

// Set column width
table.getRows().get(0).getCells().get(0).setCellWidth(70f,CellWidthType.Point);
table.getRows().get(1).getCells().get(0).setCellWidth(70f,CellWidthType.Point);
table.autoFit(AutoFitBehaviorType.Auto_Fit_To_Window);

// Insert content to cells
table.get(0, 0).addParagraph().appendText("Spire.Doc for Java");
table.get(0, 1).addParagraph().appendText("Sample text");

// Add a nested table to cell (first row, second column)
Table nestedTable = table.get(0, 1).addTable(true);
nestedTable.resetCells(4, 3);
nestedTable.autoFit(AutoFitBehaviorType.Auto_Fit_To_Contents);

// Add content to nested cells
nestedTable.get(0, 0).addParagraph().appendText("NO.");
nestedTable.get(0, 1).addParagraph().appendText("Item");
nestedTable.get(0, 2).addParagraph().appendText("Price");

nestedTable.get(1, 0).addParagraph().appendText("1");
nestedTable.get(1, 1).addParagraph().appendText("Pro Edition");
nestedTable.get(1, 2).addParagraph().appendText("$799");

nestedTable.get(2, 0).addParagraph().appendText("2");
nestedTable.get(2, 1).addParagraph().appendText("Standard Edition");
nestedTable.get(2, 2).addParagraph().appendText("$599");

nestedTable.get(3, 0).addParagraph().appendText("3");
nestedTable.get(3, 1).addParagraph().appendText("Free Edition");
nestedTable.get(3, 2).addParagraph().appendText("$0");
```

---

# Spire.Doc Table Creation
## Create and format a table in a Word document
```java
// Create the table
Table table = section.addTable(true);
table.resetCells(data.length + 1, header.length);


// Add the header row
TableRow row = table.getRows().get(0);
row.isHeader(true);
row.setHeight(20);
row.setHeightType(TableRowHeightType.Exactly);

for (int j = 0; j < row.getCells().getCount(); j++)
{
    row.getCells().get(j).getCellFormat().getShading().setBackgroundPatternColor(Color.gray);
}

for (int i = 0; i < header.length; i++)
{
    row.getCells().get(i).getCellFormat().setVerticalAlignment(VerticalAlignment.Middle);
    Paragraph p = row.getCells().get(i).addParagraph();
    p.getFormat().setHorizontalAlignment(HorizontalAlignment.Center);
    TextRange txtRange = p.appendText(header[i]);
    txtRange.getCharacterFormat().setBold(true);
}
// Add data rows
for (int r = 0; r < data.length; r++)
{
    TableRow dataRow = table.getRows().get(r + 1);
    dataRow.setHeight(25);
    dataRow.setHeightType(TableRowHeightType.Exactly);
    for (int c = 0; c < dataRow.getCells().getCount(); c++)
    {
        dataRow.getCells().get(c).getCellFormat().getShading().setBackgroundPatternColor(Color.white);
    }

    for (int c = 0; c < data[r].length; c++)
    {
        dataRow.getCells().get(c).getCellFormat().setVerticalAlignment(VerticalAlignment.Middle);
        dataRow.getCells().get(c).addParagraph().appendText(data[r][c]);
    }
}
// Apply alternating row color
for (int j = 1; j < table.getRows().getCount(); j++)
{
    if (j % 2 == 0)
    {
        TableRow row2 = table.getRows().get(j);
        for (int f = 0; f < row2.getCells().getCount(); f++)
        {
            row2.getCells().get(f).getCellFormat().getShading().setBackgroundPatternColor(new Color(173, 216, 230)/*Color.getLightBlue()*/);
        }
    }
}
```

---

# Create Table Directly in Word Document
## This code demonstrates how to create a table directly in a Word document using Spire.Doc for Java
```java
// Create a new document object
Document doc = new Document();

// Add a section to the document
Section section = doc.addSection();

// Create a table object
Table table = new Table(doc);

//Set the width of table
table.setPreferredWidth(new PreferredWidth(WidthType.Percentage, (short)100));
//Set the border of table
table.getFormat().getBorders().setBorderType(BorderStyle.Single);

//Create a table row
TableRow row = table.getRows().get(0);
row.setHeight(50.0f);

//Create a table cell
TableCell cell1 = table.getRows().get(0).getCells().get(0);
//Add a paragraph
Paragraph para1 = cell1.addParagraph();
//Append text in the paragraph
para1.appendText("Row 1, Cell 1");
//Set the horizontal alignment of paragrah
para1.getFormat().setHorizontalAlignment(HorizontalAlignment.Center);
//Set the background color of cell
cell1.getCellFormat().getShading().setBackgroundPatternColor(Color.lightGray);
//Set the vertical alignment of paragraph
cell1.getCellFormat().setVerticalAlignment(VerticalAlignment.Middle);

//Create a table cell
TableCell cell2 = table.getRows().get(0).getCells().get(1);
Paragraph para2 = cell2.addParagraph();
para2.appendText("Row 1, Cell 2");
para2.getFormat().setHorizontalAlignment(HorizontalAlignment.Center);
cell2.getCellFormat().getShading().setBackgroundPatternColor(Color.lightGray);
cell2.getCellFormat().setVerticalAlignment(VerticalAlignment.Middle);

//Add the table in the section
section.getTables().add(table);
```

---

# Spire.Doc HTML Table Creation
## Create a table in a Word document from HTML content
```java
// Create a new document object
Document document = new Document();

// Add a section to the document
Section section = document.addSection();

// Add a paragraph to the section and append the HTML content
section.addParagraph().appendHTML(htmlString);

// Dispose the document resources
document.dispose();
```

---

# Spire.Doc Vertical Table Creation
## Create a table with vertical text direction in Word document
```java
// Add a table to the section
Table table = section.addTable();

// Reset the table with 1 row and 1 column
table.resetCells(1, 1);

// Get the cell in the first row and first column
TableCell cell = table.getRows().get(0).getCells().get(0);

// Set the height of the first row to 150f (float)
table.getRows().get(0).setHeight(150f);

// Add a paragraph with text to the cell
cell.addParagraph().appendText("Draft copy in vertical style");

// Set the text direction to right-to-left rotated
cell.getCellFormat().setTextDirection(TextDirection.Right_To_Left_Rotated);

//Set the table format.
table.getFormat().setWrapTextAround(true);
table.getFormat().getPositioning().setVertRelationTo(VerticalRelation.Page);
table.getFormat().getPositioning().setHorizRelationTo(HorizontalRelation.Page);
table.getFormat().getPositioning().setHorizPosition((float)section.getPageSetup().getPageSize().getWidth() - table.getWidth());
table.getFormat().getPositioning().setVertPosition(200f);
```

---

# Spire.Doc Table Formatting
## Format merged cells in a Word document table
```java
// Create a new document object
Document document = new Document();

// Add a section to the document
Section section = document.addSection();

// Add a table to the section
Table table = section.addTable(true);
table.resetCells(4, 3);

// Create a new paragraph style
ParagraphStyle style = new ParagraphStyle(document);
style.setName("Style");
style.getCharacterFormat().setTextColor(Color.cyan);
style.getCharacterFormat().setItalic(true);
style.getCharacterFormat().setBold(true);
style.getCharacterFormat().setFontSize(13);
document.getStyles().add(style);

// Apply horizontal merge to cells in the table
table.applyHorizontalMerge(0, 0, 1);

// Apply the style to the first cell in the table
table.get(0, 0).getParagraphs().get(0).applyStyle(style.getName());
table.get(0, 0).getCellFormat().setVerticalAlignment(VerticalAlignment.Middle);
table.get(0, 0).getParagraphs().get(0).getFormat().setHorizontalAlignment(HorizontalAlignment.Center);

// Apply vertical merge to cells in the table
table.applyVerticalMerge(0, 1, 3);

// Apply the style to the second cell in the table
table.get(1, 0).getParagraphs().get(0).applyStyle(style.getName());
table.get(1, 0).getCellFormat().setVerticalAlignment(VerticalAlignment.Middle);
table.get(1, 0).getParagraphs().get(0).getFormat().setHorizontalAlignment(HorizontalAlignment.Left);

// Set the width of the second cell in the table
table.get(1, 0).setCellWidth(20, CellWidthType.Percentage);
```

---

# spire.doc diagonal border properties
## get diagonal border properties from table cells in a Word document
```java
// Get the first section of the document
Section section = document.getSections().get(0);

// Get the first table in the section
Table table = section.getTables().get(0);

// Get the diagonal down border style, line width, and color of the cell at (0, 0)
BorderStyle cellBorderStyle_down = table.get(0, 0).getCellFormat().getBorders().getDiagonalDown().getBorderType();
float cellBorderLineWidth_down = table.get(0, 0).getCellFormat().getBorders().getDiagonalDown().getLineWidth();
Color cellBorderColor_down = table.get(0, 0).getCellFormat().getBorders().getDiagonalDown().getColor();

// Get the diagonal up border style, line width, and color of the cell at (3, 2)
BorderStyle cellBorderStyle_up = table.get(3, 2).getCellFormat().getBorders().getDiagonalUp().getBorderType();
float cellBorderLineWidth_up = table.get(3, 2).getCellFormat().getBorders().getDiagonalUp().getLineWidth();
Color cellBorderColor_up = table.get(3, 2).getCellFormat().getBorders().getDiagonalUp().getColor();
```

---

# Spire.Doc Table Index Retrieval
## Get table, row, and cell indices from a Word document
```java
// Get the first section of the document
Section section = doc.getSections().get(0);

// Get the first table in the section
Table table = section.getTables().get(0);

// Get the collection of tables in the section
TableCollection collections = section.getTables();

// Get the index of the table in the collection
int tableIndex = collections.indexOf(table);

// Get the last row in the table
TableRow row = table.getLastRow();

// Get the index of the row
int rowIndex = row.getRowIndex();

// Get the last cell in the row
TableCell cell = (TableCell) row.getLastChild();

// Get the index of the cell
int cellIndex = cell.getCellIndex();
```

---

# Spire.Doc Table Position
## Get table positioning information from a Word document
```java
// Get the first section of the document
Section section = document.getSections().get(0);

// Get the first table in the section
Table table = section.getTables().get(0);

// Check if text wrapping is enabled for the table
if (table.getFormat().getWrapTextAround()) {
    // Get the positioning information for the table
    TablePositioning position = table.getFormat().getPositioning();

    // Get horizontal positioning details
    String horizPosition = "Position: " + position.getHorizPosition() + " pt";
    String horizAbsPosition = "Absolute Position: " + position.getHorizPositionAbs() +
            ", Relative to: " + position.getHorizRelationTo();

    // Get vertical positioning details
    String vertPosition = "Position: " + position.getVertPosition() + " pt";
    String vertAbsPosition = "Absolute Position: " + position.getVertPositionAbs() +
            ", Relative to: " + position.getVertRelationTo();

    // Get distance from surrounding text
    String distance = "Top: " + position.getDistanceFromTop() + " pt, Left: " + 
            position.getDistanceFromLeft() + " pt\nBottom: " + 
            position.getDistanceFromBottom() + " pt, Right: " + 
            position.getDistanceFromRight() + " pt";
}
```

---

# Table Cell Operations
## Merge and split table cells in a document
```java
// Get the first section of the document
Section section = document.getSections().get(0);

// Get the first table in the section
Table table = section.getTables().get(0);

// Apply horizontal merge to cells in the specified range (row 6, column 2 to row 6, column 3)
table.applyHorizontalMerge(6, 2, 3);

// Apply vertical merge to cells in the specified range (row 2, column 4 to row 2, column 5)
table.applyVerticalMerge(2, 4, 5);

// Split a cell into multiple cells at the specified position (row 8, column 3), creating a 2x2 cell grid
table.getRows().get(8).getCells().get(3).splitCell(2, 2);
```

---

# Spire.Doc Table Formatting
## Modify table, row, and cell formats in Word documents
```java
// Method to modify the table format
private static void ModifyTableFormat(Table table)
{
    // Set the preferred width of the table
    table.setPreferredWidth(new PreferredWidth(WidthType.Twip, (short)6000));

    // Apply a predefined table style to the table
    table.applyStyle(DefaultTableStyle.Table_3_Deffects_2);

    // Set padding for all sides of the table
    table.getFormat().getPaddings().setAll(5f);

    // Set the title and description for the table
    table.setTitle("Spire.Doc for Java");
    table.setTableDescription("Spire.Doc for Java is a professional Word Java library");
}

// Method to modify the row format
private static void ModifyRowFormat(Table table)
{
    //Set cell spacing
    table.getFormat().setCellSpacing(2f);

    //Set row height
    table.getRows().get(1).setHeightType(TableRowHeightType.Exactly);
    table.getRows().get(1).setHeight(20f);

    //Set background color
    for (int i = 0; i < table.getRows().get(2).getCells().getCount(); i++)
    {
        table.getRows().get(2).getCells().get(i).getCellFormat().getShading().setBackgroundPatternColor(Color.gray);
    }
}

// Method to modify the cell format
private static void ModifyCellFormat(Table table)
{
    //Set alignment
    table.getRows().get(0).getCells().get(0).getCellFormat().setVerticalAlignment(VerticalAlignment.Middle);
    table.getRows().get(0).getCells().get(0).getParagraphs().get(0).getFormat().setHorizontalAlignment(HorizontalAlignment.Center);

    //Set background color
    table.getRows().get(1).getCells().get(0).getCellFormat().getShading().setBackgroundPatternColor(Color.gray);

    //Set cell border
    table.getRows().get(2).getCells().get(0).getCellFormat().getBorders().setBorderType(BorderStyle.Single);
    table.getRows().get(2).getCells().get(0).getCellFormat().getBorders().setLineWidth(1f);
    table.getRows().get(2).getCells().get(0).getCellFormat().getBorders().getLeft().setColor(Color.red);
    table.getRows().get(2).getCells().get(0).getCellFormat().getBorders().getRight().setColor(Color.red);
    table.getRows().get(2).getCells().get(0).getCellFormat().getBorders().getTop().setColor(Color.red);
    table.getRows().get(2).getCells().get(0).getCellFormat().getBorders().getBottom().setColor(Color.red);

    //Set text direction
    table.getRows().get(3).getCells().get(0).getCellFormat().setTextDirection(TextDirection.Right_To_Left);
}
```

---

# Spire.Doc Table Page Break Prevention
## Prevent page breaks in Word table paragraphs
```java
// Get the first table in the first section of the document
Table table = document.getSections().get(0).getTables().get(0);

// Iterate over each row in the table
for (TableRow row : (Iterable<TableRow>) table.getRows()) {
    // Iterate over each cell in the row
    for (TableCell cell : (Iterable<TableCell>) row.getCells()) {
        // Iterate over each paragraph in the cell
        for (Paragraph p : (Iterable<Paragraph>) cell.getParagraphs()) {
            // Set the keep follow property to true, preventing page breaks within the paragraph
            p.getFormat().setKeepFollow(true);
        }
    }
}
```

---

# Spire.Doc Table Removal
## Remove a table from a Word document
```java
// Get the first section of the document and remove the first table in it
doc.getSections().get(0).getTables().removeAt(0);
```

---

# Spire.Doc Table Header Row Repetition
## Create a table with header rows that repeat on each page
```java
//Create a table width default borders
Table table = section.addTable(true);
//Set table with to 100%
PreferredWidth width = new PreferredWidth(WidthType.Percentage, (short)100);
table.setPreferredWidth(width);

//Add a new row
TableRow row = table.addRow();
//Set the row as a table header
row.isHeader(true);
//Add a new cell for row
TableCell cell = row.addCell();

//Set the backcolor of row
for (int j = 0; j < row.getCells().getCount(); j++)
{
    row.getCells().get(j).getCellFormat().getShading().setBackgroundPatternColor(Color.lightGray);
}
cell.setCellWidth(100, CellWidthType.Percentage);
//Add a paragraph for cell to put some data
Paragraph parapraph = cell.addParagraph();
//Add text
parapraph.appendText("Row Header 1");
//Set paragraph horizontal center alignment
parapraph.getFormat().setHorizontalAlignment(HorizontalAlignment.Center);

row = table.addRow(false, 1);
row.isHeader(true);
for (int j = 0; j < row.getCells().getCount(); j++)
{
    row.getCells().get(j).getCellFormat().getShading().setBackgroundPatternColor(Color.orange);
}
//Set row height
row.setHeight(30f);
cell = row.getCells().get(0);
cell.setCellWidth(100, CellWidthType.Percentage);
//Set cell vertical middle alignment
cell.getCellFormat().setVerticalAlignment(VerticalAlignment.Middle);
//Add a paragraph for cell to put some data
parapraph = cell.addParagraph();
//Add text
parapraph.appendText("Row Header 2");
parapraph.getFormat().setHorizontalAlignment(HorizontalAlignment.Center);

//Add many common rows
for (int i = 0; i < 70; i++)
{
    row = table.addRow(false, 2);
    cell = row.getCells().get(0);
    //Set cell width to 50% of table width
    cell.setCellWidth(50f, CellWidthType.Percentage);
    cell.addParagraph().appendText("Column 1 Text");
    cell = row.getCells().get(1);
    cell.setCellWidth(50f, CellWidthType.Percentage);
    cell.addParagraph().appendText("Column 2 Text");
}
//Set cell backcolor
for (int j = 1; j < table.getRows().getCount(); j++)
{
    if (j % 2 == 0)
    {
        TableRow row2 = table.getRows().get(j);
        for (int f = 0; f < row2.getCells().getCount(); f++)
        {
            row2.getCells().get(f).getCellFormat().getShading().setBackgroundPatternColor(Color.PINK);
        }
    }
}
```

---

# Spire.Doc Table Text Replacement
## Replace text in table using regular expressions and exact text matching
```java
// Create a new document object
Document doc = new Document();

// Get the first section of the document
Section section = doc.getSections().get(0);

// Get the first table in the section
Table table = section.getTables().get(0);

// Define a regular expression pattern to match text within curly braces {}
Pattern regex = Pattern.compile("\\{[^\\}]+\\}", 0);

// Replace the matched text with "E-iceblue" in the table
table.replace(regex, "E-iceblue");

// Replace the exact text "Beijing" with "Component" in the table,
// ignoring case and allowing partial word matching
table.replace("Beijing", "Component", false, true);
```

---

# Spire.Doc Table Column Width
## Set the width of table cells in a Word document
```java
// Get the first section of the document
Section section = document.getSections().get(0);

// Get the first table in the section
Table table = section.getTables().get(0);

// Loop through each row in the table
for (int i = 0; i < table.getRows().getCount(); i++) {
    // Set the width type of the first cell in each row to point, 200 points
    table.getRows().get(i).getCells().get(0).setCellWidth(200,CellWidthType.Point);
}
```

---

# Spire.Doc Diagonal Border Setting
## Set diagonal borders for table cells in a Word document
```java
// Create a new document object
Document document = new Document();

// Add a section to the document
Section section = document.addSection();

// Add a table to the section, with autofit behavior enabled
Table table = section.addTable(true);

// Reset the number of cells in the table to 4 columns and 3 rows
table.resetCells(4, 3);

// Set the horizontal alignment of the table to center
table.getFormat().setHorizontalAlignment(RowAlignment.Center);

// Set the diagonal down border for the first cell in the first row
table.getFirstRow().getCells().get(0).getCellFormat().getBorders().getDiagonalDown().setBorderType(BorderStyle.Double);
table.getFirstRow().getCells().get(0).getCellFormat().getBorders().getDiagonalDown().setColor(Color.GREEN);
table.getFirstRow().getCells().get(0).getCellFormat().getBorders().getDiagonalDown().setLineWidth(2f);

// Set the diagonal up border for the last cell in the table
table.getLastCell().getCellFormat().getBorders().getDiagonalUp().setBorderType(BorderStyle.Single);
table.getLastCell().getCellFormat().getBorders().getDiagonalUp().setColor(Color.RED);
table.getLastCell().getCellFormat().getBorders().getDiagonalUp().setLineWidth(0.8f);
```

---

# spire.doc table positioning
## set table outside position in document header
```java
// Add a table to the header
Table table = header.addTable();

// Reset the number of cells in the table to 4 rows and 2 columns
table.resetCells(4, 2);

// Set the table's wrap text around property to true
table.getFormat().setWrapTextAround(true);

// Set the table's horizontal absolute positioning to "Outside" the text area
table.getFormat().getPositioning().setHorizPositionAbs(HorizontalPosition.Outside);

// Set the table's vertical positioning relative to the margin and set the vertical position to 43 points
table.getFormat().getPositioning().setVertRelationTo(VerticalRelation.Margin);
table.getFormat().getPositioning().setVertPosition(43f);
```

---

# Spire.Doc Table Styling
## Set table style and borders in Word document
```java
// Get the first section of the document
Section section = document.getSections().get(0);

// Get the first table in the section
Table table = section.getTables().get(0);

// Apply the "Colorful_List" default table style to the table
table.applyStyle(DefaultTableStyle.Colorful_List);

// Set the right border of the table to a red hairline with a line width of 1.0F
table.getFormat().getBorders().getRight().setBorderType(BorderStyle.Hairline);
table.getFormat().getBorders().getRight().setLineWidth(1.0F);
table.getFormat().getBorders().getRight().setColor(Color.RED);

// Set the top border of the table to a green hairline with a line width of 1.0F
table.getFormat().getBorders().getTop().setBorderType(BorderStyle.Hairline);
table.getFormat().getBorders().getTop().setLineWidth(1.0F);
table.getFormat().getBorders().getTop().setColor(Color.GREEN);

// Set the left border of the table to a yellow hairline with a line width of 1.0F
table.getFormat().getBorders().getLeft().setBorderType(BorderStyle.Hairline);
table.getFormat().getBorders().getLeft().setLineWidth(1.0F);
table.getFormat().getBorders().getLeft().setColor(Color.YELLOW);

// Set the bottom border of the table to a dot-dash style
table.getFormat().getBorders().getBottom().setBorderType(BorderStyle.Dot_Dash);

// Set the vertical borders of the table to a dot style and horizontal borders to none
table.getFormat().getBorders().getVertical().setBorderType(BorderStyle.Dot);
table.getFormat().getBorders().getHorizontal().setBorderType(BorderStyle.None);
table.getFormat().getBorders().getVertical().setColor(Color.ORANGE);
```

---

# spire.doc vertical alignment
## set vertical alignment for table cells in Word document
```java
// Create a new document object
Document doc = new Document();

// Add a section to the document
Section section = doc.addSection();

// Add a table to the section, with autofit behavior enabled
Table table = section.addTable(true);

// Reset the number of cells in the table to 3 columns and 3 rows
table.resetCells(3, 3);

// Apply vertical merge for the cells in the first column, from row 0 to row 2
table.applyVerticalMerge(0, 0, 2);

// Set the vertical alignment of specific cells in the table
table.getRows().get(0).getCells().get(0).getCellFormat().setVerticalAlignment(VerticalAlignment.Middle);
table.getRows().get(0).getCells().get(1).getCellFormat().setVerticalAlignment(VerticalAlignment.Top);
table.getRows().get(0).getCells().get(2).getCellFormat().setVerticalAlignment(VerticalAlignment.Top);
table.getRows().get(1).getCells().get(1).getCellFormat().setVerticalAlignment(VerticalAlignment.Middle);
table.getRows().get(1).getCells().get(2).getCellFormat().setVerticalAlignment(VerticalAlignment.Middle);
table.getRows().get(2).getCells().get(1).getCellFormat().setVerticalAlignment(VerticalAlignment.Bottom);
table.getRows().get(2).getCells().get(2).getCellFormat().setVerticalAlignment(VerticalAlignment.Bottom);
```

---

# Spire.Doc Image Hyperlink
## Create an image hyperlink in a Word document
```java
// Create a new document object
Document doc = new Document();

// Get the first section of the document
Section section = doc.getSections().get(0);

// Add a paragraph to the section
Paragraph paragraph = section.addParagraph();

// Create a DocPicture and load an image
DocPicture picture = new DocPicture(doc);
picture.loadImage("data/spireDoc.png");

// Append a hyperlink to the paragraph with the image and set its URL and type
paragraph.appendHyperlink("https://www.e-iceblue.com/Introduce/doc-for-java.html", picture, HyperlinkType.Web_Link);
```

---

# find hyperlinks in document
## extract hyperlink fields and their text from document
```java
// Create an ArrayList to store hyperlink fields and a string to hold the extracted hyperlinks' text
ArrayList<Field> hyperlinks = new ArrayList<Field>();
String hyperlinksText = "";

// Iterate through the sections of the document
for (Section section : (Iterable<Section>) doc.getSections()) {
    // Iterate through the child objects in the section's body
    for (DocumentObject object : (Iterable<DocumentObject>) section.getBody().getChildObjects()) {
        // Check if the object is a paragraph
        if (object.getDocumentObjectType().equals(DocumentObjectType.Paragraph)) {
            Paragraph paragraph = (Paragraph) object;
            // Iterate through the child objects in the paragraph
            for (DocumentObject cObject : (Iterable<DocumentObject>) paragraph.getChildObjects()) {
                // Check if the child object is a field
                if (cObject.getDocumentObjectType().equals(DocumentObjectType.Field)) {
                    Field field = (Field) cObject;
                    // Check if the field type is a hyperlink
                    if (field.getType().equals(FieldType.Field_Hyperlink)) {
                        // Add the hyperlink field to the list
                        hyperlinks.add(field);

                        // Append the field's text to the hyperlinksText string
                        hyperlinksText += field.getFieldText() + "\r\n";
                    }
                }
            }
        }
    }
}
```

---

# Spire.Doc Hyperlinks
## Insert various types of hyperlinks in a Word document
```java
private static void insertHyperlink(Section section) throws Exception {
    // Get the first paragraph in the section or add a new one if none exists
    Paragraph paragraph = section.getParagraphs().getCount() > 0 ? section.getParagraphs().get(0) : section.addParagraph();

    // Append text to the paragraph and apply a built-in heading style
    paragraph.appendText("Spire.Doc for Java \r\n e-iceblue company Ltd. 2002-2019 All rights reserved");
    paragraph.applyStyle(BuiltinStyle.Heading_2);

    // Add a new paragraph for the "Home page" hyperlink
    paragraph = section.addParagraph();
    paragraph.appendText("Home page");
    paragraph.applyStyle(BuiltinStyle.Heading_2);

    // Add a hyperlink to the paragraph with the specified URL and display text
    paragraph = section.addParagraph();
    paragraph.appendHyperlink("www.e-iceblue.com", "www.e-iceblue.com", HyperlinkType.Web_Link);

    // Add a new paragraph for the "Contact US" hyperlink
    paragraph = section.addParagraph();
    paragraph.appendText("Contact US");
    paragraph.applyStyle(BuiltinStyle.Heading_2);

    // Add a hyperlink to the paragraph with the specified email address and display text
    paragraph = section.addParagraph();
    paragraph.appendHyperlink("mailto:support@e-iceblue.com", "support@e-iceblue.com", HyperlinkType.E_Mail_Link);

    // Add a new paragraph for the "Forum" hyperlink
    paragraph = section.addParagraph();
    paragraph.appendText("Forum");
    paragraph.applyStyle(BuiltinStyle.Heading_2);

    // Add a hyperlink to the paragraph with the specified URL and display text
    paragraph = section.addParagraph();
    paragraph.appendHyperlink("www.e-iceblue.com/forum/", "www.e-iceblue.com/forum/", HyperlinkType.Web_Link);

    // Add a new paragraph for the "Download Link" hyperlink
    paragraph = section.addParagraph();
    paragraph.appendText("Download Link");
    paragraph.applyStyle(BuiltinStyle.Heading_2);

    // Add a hyperlink to the paragraph with the specified URL and display text
    paragraph = section.addParagraph();
    paragraph.appendHyperlink("www.e-iceblue.com/Download/doc-for-java.html", "www.e-iceblue.com/Download/doc-for-java.html", HyperlinkType.Web_Link);

    // Add a new paragraph for the "Insert Link On Image" hyperlink
    paragraph = section.addParagraph();
    paragraph.appendText("Insert Link On Image");
    paragraph.applyStyle(BuiltinStyle.Heading_2);

    // Add a paragraph with an image and a hyperlink to the paragraph using the specified URL and image object
    paragraph = section.addParagraph();
    DocPicture picture = paragraph.appendPicture("data/spireDoc.png");
    paragraph.appendHyperlink("www.e-iceblue.com/Download/doc-for-java.html", picture, HyperlinkType.Web_Link);
}
```

---

# Spire.Doc Hyperlink Modification
## Code to find and modify hyperlink text in a Word document
```java
// Create an ArrayList to store hyperlink fields
ArrayList<Field> hyperlinks = new ArrayList<Field>();

// Iterate through the sections of the document
for (Section section : (Iterable<Section>) doc.getSections()) {
    // Iterate through the child objects in the section's body
    for (DocumentObject object : (Iterable<DocumentObject>) section.getBody().getChildObjects()) {
        // Check if the object is a paragraph
        if (object.getDocumentObjectType().equals(DocumentObjectType.Paragraph)) {
            Paragraph paragraph = (Paragraph) object;
            // Iterate through the child objects in the paragraph
            for (DocumentObject cObject : (Iterable<DocumentObject>) paragraph.getChildObjects()) {
                // Check if the child object is a field
                if (cObject.getDocumentObjectType().equals(DocumentObjectType.Field)) {
                    Field field = (Field) cObject;
                    // Check if the field type is a hyperlink
                    if (field.getType().equals(FieldType.Field_Hyperlink)) {
                        // Add the hyperlink field to the list
                        hyperlinks.add(field);
                    }
                }
            }
        }
    }
}

// Modify the text of the first hyperlink field in the list
hyperlinks.get(0).setFieldText("Modified Text");
```

---

# Spire.Doc Hyperlink Removal
## Remove all hyperlinks from a Word document while preserving the displayed text
```java
// Find all hyperlinks in the document and store them in an ArrayList
ArrayList<Field> hyperlinks = FindAllHyperlinks(doc);

// Iterate through the hyperlinks in reverse order and flatten them
for (int i = hyperlinks.size() - 1; i >= 0; i--) {
    FlattenHyperlinks(hyperlinks.get(i));
}

// Find all hyperlinks in the document and return them as an ArrayList
private static ArrayList<Field> FindAllHyperlinks(Document document) {
    ArrayList<Field> hyperlinks = new ArrayList<Field>();

    // Iterate through the sections of the document
    for (Section section : (Iterable<Section>) document.getSections()) {
        // Iterate through the child objects in the section's body
        for (DocumentObject object : (Iterable<DocumentObject>) section.getBody().getChildObjects()) {
            // Check if the object is a paragraph
            if (object.getDocumentObjectType().equals(DocumentObjectType.Paragraph)) {
                Paragraph paragraph = (Paragraph) object;
                // Iterate through the child objects in the paragraph
                for (DocumentObject cObject : (Iterable<DocumentObject>) paragraph.getChildObjects()) {
                    // Check if the child object is a field
                    if (cObject.getDocumentObjectType().equals(DocumentObjectType.Field)) {
                        Field field = (Field) cObject;
                        // Check if the field type is a hyperlink
                        if (field.getType().equals(FieldType.Field_Hyperlink)) {
                            // Add the hyperlink field to the list
                            hyperlinks.add(field);
                        }
                    }
                }
            }
        }
    }
    return hyperlinks;
}

// Flatten a hyperlink by removing its field structure and keeping only the displayed text
private static void FlattenHyperlinks(Field field) {
    // Get the indices of the field and its related objects in the document structure
    int ownerParaIndex = field.getOwnerParagraph().getOwnerTextBody().getChildObjects().indexOf(field.getOwnerParagraph());
    int fieldIndex = field.getOwnerParagraph().getChildObjects().indexOf(field);
    int sepOwnerParaIndex = field.getSeparator().getOwnerParagraph().getOwnerTextBody().getChildObjects().indexOf(field.getSeparator().getOwnerParagraph());
    int sepIndex = field.getSeparator().getOwnerParagraph().getChildObjects().indexOf(field.getSeparator());
    int endIndex = field.getEnd().getOwnerParagraph().getChildObjects().indexOf(field.getEnd());
    int endOwnerParaIndex = field.getEnd().getOwnerParagraph().getOwnerTextBody().getChildObjects().indexOf(field.getEnd().getOwnerParagraph());

    // Format the text between the separator and the end of the field result
    FormatFieldResultText(field.getSeparator().getOwnerParagraph().getOwnerTextBody(), sepOwnerParaIndex, endOwnerParaIndex, sepIndex, endIndex);

    // Remove the end object of the field
    field.getEnd().getOwnerParagraph().getChildObjects().removeAt(endIndex);

    // Iterate through the objects to be removed and remove them from the document structure
    for (int i = sepOwnerParaIndex; i >= ownerParaIndex; i--) {
        if (i == sepOwnerParaIndex && i == ownerParaIndex) {
            for (int j = sepIndex; j >= fieldIndex; j--) {
                field.getOwnerParagraph().getChildObjects().removeAt(j);
            }
        } else if (i == ownerParaIndex) {
            for (int j = field.getOwnerParagraph().getChildObjects().getCount() - 1; j >= fieldIndex; j--) {
                field.getOwnerParagraph().getChildObjects().removeAt(j);
            }
        } else if (i == sepOwnerParaIndex) {
            for (int j = sepIndex; j >= 0; j--) {
                field.getOwnerParagraph().getChildObjects().removeAt(j);
            }
        } else {
            field.getOwnerParagraph().getOwnerTextBody().getChildObjects().removeAt(i);
        }
    }
}

// Format the text within a field result based on specified indices
private static void FormatFieldResultText(Body ownerBody, int sepOwnerParaIndex, int endOwnerParaIndex, int sepIndex, int endIndex) {
    for (int i = sepOwnerParaIndex; i <= endOwnerParaIndex; i++) {
        Paragraph para = (Paragraph) ownerBody.getChildObjects().get(i);
        if (i == sepOwnerParaIndex && i == endOwnerParaIndex) {
            // Iterate through the child objects in the paragraph between the separator and the end index
            for (int j = sepIndex + 1; j < endIndex; j++) {
                FormatText((TextRange) para.getChildObjects().get(j));
            }
        } else if (i == sepOwnerParaIndex) {
            // Iterate through the remaining child objects in the paragraph starting from the separator index
            for (int j = sepIndex + 1; j < para.getChildObjects().getCount(); j++) {
                FormatText((TextRange) para.getChildObjects().get(j));
            }
        } else if (i == endOwnerParaIndex) {
            // Iterate through the child objects in the paragraph up to the end index
            for (int j = 0; j < endIndex; j++) {
                FormatText((TextRange) para.getChildObjects().get(j));
            }
        } else {
            // Iterate through all the child objects in the paragraph
            for (int j = 0; j < para.getChildObjects().getCount(); j++) {
                FormatText((TextRange) para.getChildObjects().get(j));
            }
        }
    }
}

// Format the text range by setting the font color to black and removing underline
private static void FormatText(TextRange tr) {
    tr.getCharacterFormat().setTextColor(Color.black);
    tr.getCharacterFormat().setUnderlineStyle(UnderlineStyle.None);
}
```

---

# Spire.Doc hyperlink formatting
## Set different formatting options for hyperlinks in Word documents
```java
// Create a paragraph for the regular link
Paragraph para1 = section.addParagraph();
para1.appendText("Regular Link: ");

// Append a hyperlink to the paragraph with the specified URL and display text
TextRange txtRange1 = para1.appendHyperlink("www.e-iceblue.com", "www.e-iceblue.com", HyperlinkType.Web_Link);
txtRange1.getCharacterFormat().setFontName("Times New Roman");
txtRange1.getCharacterFormat().setFontSize(12f);

// Insert a blank paragraph
Paragraph blankPara1 = section.addParagraph();

// Create a paragraph for the link with changed color
Paragraph para2 = section.addParagraph();
para2.appendText("Change Color: ");

// Append a hyperlink to the paragraph with the specified URL and display text
TextRange txtRange2 = para2.appendHyperlink("www.e-iceblue.com", "www.e-iceblue.com", HyperlinkType.Web_Link);
txtRange2.getCharacterFormat().setFontName("Times New Roman");
txtRange2.getCharacterFormat().setFontSize(12f);
txtRange2.getCharacterFormat().setTextColor(Color.red);

// Insert a blank paragraph
Paragraph blankPara2 = section.addParagraph();

// Create a paragraph for the link with removed underline
Paragraph para3 = section.addParagraph();
para3.appendText("Remove Underline: ");

// Append a hyperlink to the paragraph with the specified URL and display text
TextRange txtRange3 = para3.appendHyperlink("www.e-iceblue.com", "www.e-iceblue.com", HyperlinkType.Web_Link);
txtRange3.getCharacterFormat().setFontName("Times New Roman");
txtRange3.getCharacterFormat().setFontSize(12f);
txtRange3.getCharacterFormat().setUnderlineStyle(UnderlineStyle.None);
```

---

# Spire.Doc Digital Signature
## Add digital signature to document
```java
// Create document object
Document doc = new Document();

// Save the document with digital signature
doc.saveToFile(output, FileFormat.Docx, certificatePath, securePwd);
```

---

# Spire.Doc Digital Signature Check
## Check if a Word document has a digital signature
```java
// Determine if a document has a digital signature
boolean value = Document.hasDigitalSignature("data/Sample.docx");
if(value) {
    // This Word document has a digital signature
} else {
    // This Word document has not a digital signature
}
```

---

# Spire.Doc Document Protection Password Check
## Check if a password matches the protection password of a document
```java
// Create a new document object
Document document = new Document();

// Load the document from the specified input file
document.loadFromFile(input);

// Check if the specified password matches the protection password of the document
boolean checkResult = document.checkProtectionPassWord(password);
```

---

# Document Decryption
## Decrypt a password-protected Word document
```java
// Create a new document object
Document document = new Document();

// Load the document with the specified format and password
document.loadFromFile(inputFile, FileFormat.Docx, "E-iceblue");
```

---

# Document Encryption
## Encrypt a Word document with password protection
```java
// Create a new document object
Document document = new Document();

// Load the document from the specified input file
document.loadFromFile(inputFile);

// Encrypt the document with the specified password
document.encrypt("E-iceblue");

// Save the encrypted document to the specified output file in DOCX format
document.saveToFile(outputFile, FileFormat.Docx);

// Dispose the document resources
document.dispose();
```

---

# Document Section Locking
## Lock specified sections in a document while allowing others to remain editable
```java
// Protect the document by allowing only form fields and using the password "123"
document.protect(ProtectionType.Allow_Only_Form_Fields, "123");

// Disable form protection for section 2
s2.setProtectForm(false);
```

---

# Spire.Doc Remove Editable Range
## Remove PermissionStart and PermissionEnd objects from Word document
```java
// Create a new document object
Document document = new Document();

// Iterate through each section in the document
for (int j = 0; j < document.getSections().getCount(); j++) {
    Section section = document.getSections().get(j);

    // Iterate through each paragraph in the section
    for (int k = 0; k < section.getParagraphs().getCount(); k++) {
        Paragraph paragraph = section.getParagraphs().get(k);

        // Iterate through each child object in the paragraph
        for (int i = 0; i < paragraph.getChildObjects().getCount(); ) {
            DocumentObject obj = paragraph.getChildObjects().get(i);

            // Remove PermissionStart and PermissionEnd objects from the paragraph
            if (obj instanceof PermissionStart || obj instanceof PermissionEnd) {
                paragraph.getChildObjects().remove(obj);
            } else {
                i++;
            }
        }
    }
}
```

---

# Spire.Doc Document Protection
## Remove read-only restriction from Word document
```java
// Create a new document object
Document document = new Document();

// Remove any read-only restrictions from the document by setting No_Protection as the protection type
document.protect(ProtectionType.No_Protection);
```

---

# Document Editable Range Protection
## Set editable ranges and protect document with password
```java
// Create a PermissionStart object with the specified permission ID
PermissionStart start = new PermissionStart(document, permissionId);

// Create a PermissionEnd object with the specified permission ID
PermissionEnd end = new PermissionEnd(document, permissionId);

// Insert the PermissionStart object at the beginning of the child objects of the first paragraph in the section
section.getParagraphs().get(0).getChildObjects().insert(0, start);

// Add the PermissionEnd object at the end of the child objects of the first paragraph in the section
section.getParagraphs().get(0).getChildObjects().add(end);

// Protect the document by allowing only reading and using the specified password
document.protect(ProtectionType.Allow_Only_Reading, password);
```

---

# Document Protection Type
## Specify protection type for a document
```java
// Protect the document by allowing only reading and using the password "123456"
document.protect(ProtectionType.Allow_Only_Reading, "123456");
```

---

# Spire.Doc Word to PDF Encryption
## Convert Word document to encrypted PDF
```java
// Create a new document object
Document document = new Document();

// Load the document
document.loadFromFile("input.docx");

// Create a ToPdfParameterList object for converting to PDF
ToPdfParameterList toPdf = new ToPdfParameterList();

// Encrypt the PDF with owner password, user password, default permissions, and 128-bit encryption key size
toPdf.getPdfSecurity().encrypt("e-iceblue", "test", PdfPermissionsFlags.Default, PdfEncryptionKeySize.Key_128_Bit);

// Save the converted PDF with encryption
document.saveToFile("output.pdf", toPdf);

// Dispose the document resources
document.dispose();
```

---

# Spire.Doc TC Field Addition
## Add a Table of Contents Entry field to a Word document
```java
// Create a new document object
Document document = new Document();

// Add a section to the document
Section section = document.addSection();

// Add a paragraph to the section
Paragraph paragraph = section.addParagraph();

// Append a TC field with the specified entry text to the paragraph
Field field = paragraph.appendField("TC", FieldType.Field_TOC_Entry);
field.setCode("TC " + "\"Entry Text\"" + " \\f" + " t");
```

---

# Spire.Doc Field Conversion
## Convert form fields to body text in Word documents
```java
// Iterate through each form field in the first section of the document
for (FormField field : (Iterable<FormField>) sourceDocument.getSections().get(0).getBody().getFormFields()) {

    // Check if the form field type is Field_Form_Text_Input
    if (field.getType().equals(FieldType.Field_Form_Text_Input)) {
        
        // Get the owner paragraph of the form field
        Paragraph paragraph = field.getOwnerParagraph();
        
        int startIndex = 0;
        int endIndex = 0;

        // Create a new text range with the content of the paragraph
        TextRange textRange = new TextRange(sourceDocument);
        textRange.setText(paragraph.getText());

        // Find the start and end index of the bookmark tags in the paragraph's child objects
        for (DocumentObject obj : (Iterable<DocumentObject>)paragraph.getChildObjects()) {
            if (obj.getDocumentObjectType().equals(DocumentObjectType.Bookmark_Start)) {
                startIndex = paragraph.getChildObjects().indexOf(obj);
            }
            if (obj.getDocumentObjectType().equals(DocumentObjectType.Bookmark_End)) {
                endIndex = paragraph.getChildObjects().indexOf(obj);
            }
        }

        // Remove any form fields or other objects between the bookmark tags
        for (int i = endIndex; i > startIndex; i--) {
            if (paragraph.getChildObjects().get(i) instanceof TextFormField) {
                TextFormField textFormField = (TextFormField) paragraph.getChildObjects().get(i);

                // Remove the text form field
                paragraph.getChildObjects().remove(textFormField);
            } else {
                // Remove the object at the specified index
                paragraph.getChildObjects().removeAt(i);
            }
        }

        // Insert the text range at the start index of the bookmark tags
        paragraph.getChildObjects().insert(startIndex, textRange);
        
        // Exit the loop after processing the first form field
        break;
    }
}
```

---

# Converting IF Fields to Text
## Converting IF fields to text while preserving formatting in a Word document
```java
// Get the collection of fields in the document
FieldCollection fields = document.getFields();

// Iterate through each field in the collection
for (int i = 0; i < fields.getCount(); i++) {
	Field field = fields.get(i);
	
	// Check if the field type is Field_If
	if (field.getType().equals(FieldType.Field_If)) {
		
		// Cast the field to TextRange to access its properties
		TextRange original = (TextRange) field;
		
		// Get the field text
		String text = field.getFieldText();
		
		// Create a new TextRange with the same text as the field
		TextRange textRange = new TextRange(document);
		textRange.setText(text);
		
		// Set the character format of the new TextRange to match the original field's font name and size
		textRange.getCharacterFormat().setFontName(original.getCharacterFormat().getFontName());
		textRange.getCharacterFormat().setFontSize(original.getCharacterFormat().getFontSize());

		// Get the owner paragraph of the field
		Paragraph par = field.getOwnerParagraph();
		
		// Get the index of the field within the child objects of the paragraph
		int index = par.getChildObjects().indexOf(field);
		
		// Remove the field from the paragraph's child objects
		par.getChildObjects().removeAt(index);
		
		// Insert the new TextRange at the same index within the paragraph's child objects
		par.getChildObjects().insert(index, textRange);
	}
}
```

---

# Spire.Doc Cross Reference Creation
## Create a cross-reference field that points to a bookmark in a Word document
```java
// Add a section to the document
Section section = document.addSection();

// Add a paragraph to the section
Paragraph paragraph = section.addParagraph();

// Append a bookmark start tag with the specified name "MyBookmark" to the paragraph
paragraph.appendBookmarkStart("MyBookmark");

// Append the text "Text inside a bookmark" to the paragraph
paragraph.appendText("Text inside a bookmark");

// Append a bookmark end tag with the specified name "MyBookmark" to the paragraph
paragraph.appendBookmarkEnd("MyBookmark");

// Add line breaks to the paragraph (repeated 4 times)
for (int i = 0; i < 4; i++) {
    paragraph.appendBreak(BreakType.Line_Break);
}

// Create a new field object
Field field = new Field(document);

// Set the field type to Field_Ref
field.setType(FieldType.Field_Ref);

// Set the field code to reference the bookmark named "MyBookmark" and include page number and hyperlink
field.setCode("REF MyBookmark \\p \\h");

// Add a new paragraph to the section
paragraph = section.addParagraph();

// Append the text "For more information, see " to the paragraph
paragraph.appendText("For more information, see ");

// Add the field as a child object of the paragraph
paragraph.getChildObjects().add(field);

// Add a field separator mark after the field
FieldMark fieldSeparator = new FieldMark(document, FieldMarkType.Field_Separator);
paragraph.getChildObjects().add(fieldSeparator);

// Add a text range object with the text "above" after the field separator mark
TextRange tr = new TextRange(document);
tr.setText("above");
paragraph.getChildObjects().add(tr);

// Add a field end mark after the text range
FieldMark fieldEnd = new FieldMark(document, FieldMarkType.Field_End);
paragraph.getChildObjects().add(fieldEnd);
```

---

# Spire.Doc Form Field Creation
## Create form fields including text input, dropdown and checkbox in a Word document
```java
// Create a new document object
Document document = new Document();

// Add a section to the document
Section section = document.addSection();

// Create field label style
ParagraphStyle formFieldLabelStyle = new ParagraphStyle(section.getDocument());
formFieldLabelStyle.setName("formFieldLabel");
formFieldLabelStyle.getCharacterFormat().setFontSize(12);
formFieldLabelStyle.getCharacterFormat().setFontName("Arial");
formFieldLabelStyle.getParagraphFormat().setHorizontalAlignment(HorizontalAlignment.Right);
section.getDocument().getStyles().add(formFieldLabelStyle);

// Add table for form fields
Table table = section.addTable();
table.setDefaultColumnsNumber(2);
table.setDefaultRowHeight(20);

// Add text input field
TableRow fieldRow1 = table.addRow(false);
fieldRow1.getCells().get(0).getCellFormat().setVerticalAlignment(VerticalAlignment.Middle);
Paragraph labelParagraph1 = fieldRow1.getCells().get(0).addParagraph();
labelParagraph1.appendText("Name:");
labelParagraph1.applyStyle(formFieldLabelStyle.getName());

fieldRow1.getCells().get(1).getCellFormat().setVerticalAlignment(VerticalAlignment.Middle);
Paragraph fieldParagraph1 = fieldRow1.getCells().get(1).addParagraph();
TextFormField field1 = (TextFormField) fieldParagraph1.appendField("name", FieldType.Field_Form_Text_Input);
field1.setDefaultText("");
field1.setText("");

// Add dropdown field
TableRow fieldRow2 = table.addRow(false);
fieldRow2.getCells().get(0).getCellFormat().setVerticalAlignment(VerticalAlignment.Middle);
Paragraph labelParagraph2 = fieldRow2.getCells().get(0).addParagraph();
labelParagraph2.appendText("Country:");
labelParagraph2.applyStyle(formFieldLabelStyle.getName());

fieldRow2.getCells().get(1).getCellFormat().setVerticalAlignment(VerticalAlignment.Middle);
Paragraph fieldParagraph2 = fieldRow2.getCells().get(1).addParagraph();
DropDownFormField list = (DropDownFormField) fieldParagraph2.appendField("country", FieldType.Field_Form_Drop_Down);
list.getDropDownItems().add("USA");
list.getDropDownItems().add("UK");
list.getDropDownItems().add("China");

// Add checkbox field
TableRow fieldRow3 = table.addRow(false);
fieldRow3.getCells().get(0).getCellFormat().setVerticalAlignment(VerticalAlignment.Middle);
Paragraph labelParagraph3 = fieldRow3.getCells().get(0).addParagraph();
labelParagraph3.appendText("Subscribe:");
labelParagraph3.applyStyle(formFieldLabelStyle.getName());

fieldRow3.getCells().get(1).getCellFormat().setVerticalAlignment(VerticalAlignment.Middle);
Paragraph fieldParagraph3 = fieldRow3.getCells().get(1).addParagraph();
fieldParagraph3.appendField("subscribe", FieldType.Field_Form_Check_Box);
```

---

# Spire.Doc IF Field Creation
## Create an IF field in a Word document
```java
// Create a new IF field
IfField ifField = new IfField(document);

// Set the field type to Field_If
ifField.setType(FieldType.Field_If);

// Set the field code for the IF field
ifField.setCode("IF");

// Add the IF field to the paragraph
paragraph.getItems().add(ifField);

// Append other text and fields to the paragraph
paragraph.appendField("Count", FieldType.Field_Merge_Field);
paragraph.appendText(" > ");
paragraph.appendText("\"100\" ");
paragraph.appendText("\"Thanks\" ");
paragraph.appendText("\"The minimum order is 100 units\"");

// Create a field end mark and add it to the paragraph
IParagraphBase endPara = document.createParagraphItem(ParagraphItemType.Field_Mark);
FieldMark end = (FieldMark)endPara;
end.setType(FieldMarkType.Field_End);
paragraph.getItems().add(end);

// Set the field end mark for the IF field
ifField.setEnd(end);
```

---

# Spire.Doc Form Field Filling
## Fill form fields in a document with data from XML
```java
// Iterate through each form field in the first section's body
for (int i = 0; i < document.getSections().get(0).getBody().getFormFields().getCount(); i++) {
    // Get the current form field
    FormField field = document.getSections().get(0).getBody().getFormFields().get(i);
    
    // Get the name of the form field
    String name = field.getName();
    
    // Iterate through each node in the XML child nodes
    for (int j = 0; j < nodeList.getLength(); j++) {
        // Get the current node
        Node node = nodeList.item(j);
        
        // Compare the form field name with the XML node name (case-insensitive)
        if (name.toLowerCase().trim().equals(node.getNodeName().toLowerCase().trim())) {
            // Get the value from the XML node
            String value = node.getTextContent();
            
            // Handle different types of form fields based on their type
            switch (field.getType()) {
                case Field_Form_Text_Input:
                    // Set the text value for a text input form field
                    field.setText(value);
                    break;
                
                case Field_Form_Drop_Down:
                    // Cast the form field to a DropDownFormField
                    DropDownFormField combox = (DropDownFormField) field;
                    
                    // Iterate through each drop-down item and find the matching value
                    for (int m = 0; m < combox.getDropDownItems().getCount(); m++) {
                        if (combox.getDropDownItems().get(m).getText().equals(value)) {
                            // Set the selected index of the drop-down field to the matching item
                            combox.setDropDownSelectedIndex(m);
                            break;
                        }
                        
                        // Special case for "country" field with "Others" option
                        if ("country".equals(field.getName()) && "Others".equals(combox.getDropDownItems().get(m).getText())) {
                            combox.setDropDownSelectedIndex(m);
                        }
                    }
                    break;
                
                case Field_Form_Check_Box:
                    // Convert the value to boolean and set the checked state for a checkbox form field
                    if (Boolean.parseBoolean(value)) {
                        CheckBoxFormField checkBox = (CheckBoxFormField) field;
                        checkBox.setChecked(true);
                    }
                    break;
            }
            
            // Break out of the loop since the matching form field has been found
            break;
        }
    }
}
```

---

# Spire.Doc Form Field Properties
## Modify form field text and character format properties
```java
// Get the first section of the document
Section section = document.getSections().get(0);

// Get the second form field from the body of the section
FormField formField = section.getBody().getFormFields().get(1);

// Check if the form field type is Field_Form_Text_Input
if (formField.getType().equals(FieldType.Field_Form_Text_Input)) {
    // Set the text of the form field to a specific value
    formField.setText("My name is " + formField.getName());
    
    // Modify the character format properties of the form field
    formField.getCharacterFormat().setTextColor(Color.red);
    formField.getCharacterFormat().setItalic(true);
}
```

---

# Extract Field Text from Document
## This code demonstrates how to extract text from fields in a document using Spire.Doc for Java
```java
// Create a StringBuilder object to store the field texts
StringBuilder sb = new StringBuilder();

// Get the collection of fields in the document
FieldCollection fields = document.getFields();

// Iterate through each field in the collection
for (Field field : (Iterable<Field>) fields) {
    // Get the text of the field
    String fieldText = field.getFieldText();
    
    // Append the field text to the StringBuilder object with additional formatting
    sb.append("The field text is \"" + fieldText + "\".\r\n");
}
```

---

# Spire.Doc Form Field Access
## Get form field by name and access its properties
```java
// Get the first section of the document
Section section = document.getSections().get(0);

// Get the form field by its name from the body of the section
FormField formField = section.getBody().getFormFields().get("email");

// Access the name and type of the form field
String fieldName = formField.getName();
String fieldType = formField.getFormFieldType().toString();
```

---

# Spire.Doc Form Fields Collection
## Get form fields collection from a document
```java
// Get the first section of the document
Section section = document.getSections().get(0);

// Get the collection of form fields from the body of the section
FormFieldCollection formFields = section.getBody().getFormFields();

// Get the count of form fields in the collection
int formFieldCount = formFields.getCount();
```

---

# Spire.Doc Merge Field Names
## Get merge field names from a document

```java
// Get the array of merge field names from the mail merge of the document
String[] fieldNames = document.getMailMerge().getMergeFieldNames();

// Get the count of merge fields in the document
int fieldCount = fieldNames.length;

// Iterate through each merge field name
for (String name : fieldNames) {
    // Process each field name
}
```

---

# Java Ask Field Handler
## Handle Ask fields in Word documents using custom field update handler
```java
// Create a custom UpdateFieldsHandler to handle AskFieldEventArgs during field update
static class HandleAskFieldex extends UpdateFieldsHandler {
    public void invoke(Object sender, IFieldsEventArgs args) {
        // Check if the event arguments are of type AskFieldEventArgs
        if (args instanceof AskFieldEventArgs) {
            AskFieldEventArgs askArgs = (AskFieldEventArgs) args;

            // Handle specific bookmark names and set appropriate response texts
            if (askArgs.getBookmarkName().equals("1")) {
                askArgs.setResponseText("Thank you. This is my first time to come to a Chinese restaurant. Could you " +
                        "tell me the different features of Chinese food?");
            }

            if (askArgs.getBookmarkName().equals("2")) {
                askArgs.setResponseText("No more, thank you. I'm quite full.");
            }
        }
    }
}

// Set the custom UpdateFieldsHandler for handling AskFieldEventArgs during field update
doc.UpdateFields = new HandleAskFieldex();

// Enable field updating during document saving
doc.isUpdateFields(true);
```

---

# Spire.Doc Address Block Field Insertion
## This code demonstrates how to insert an address block field into a Word document using Spire.Doc for Java.
```java
// Get the first section of the document
Section section = document.getSections().get(0);

// Add a new paragraph to the section
Paragraph par = section.addParagraph();

// Append an address block field to the paragraph
Field field = par.appendField("ADDRESSBLOCK", FieldType.Field_Address_Block);

// Set the code for the address block field
field.setCode("ADDRESSBLOCK \\c 1 \\d \\e Test2 \\f Test3 \\l \"Test 4\"");

// Enable updating fields in the document
document.isUpdateFields(true);
```

---

# Spire.Doc Advance Field Insertion
## Insert and configure an advance field in a Word document
```java
// Get the first section of the document
Section section = document.getSections().get(0);

// Add a new paragraph to the section
Paragraph par = section.addParagraph();

// Append an advance field to the paragraph
Field field = par.appendField("Field", FieldType.Field_Advance);

// Set the code for the advance field with specified parameters
field.setCode("ADVANCE \\d 10 \\l 10 \\r 10 \\u 0 \\x 100 \\y 100");

// Enable updating fields in the document
document.isUpdateFields(true);
```

---

# Spire.Doc Merge Field Insertion
## Insert a merge field into a Word document
```java
// Create a new empty document
Document document = new Document();

// Get the first section of the document
Section section = document.getSections().get(0);

// Add a new paragraph to the section
Paragraph par = section.addParagraph();

// Append a merge field to the paragraph with the specified field name
MergeField field = (MergeField) par.appendField("MyFieldName", FieldType.Field_Merge_Field);

// Enable updating fields in the document
document.isUpdateFields(true);
```

---

# Spire.Doc Field Insertion
## Insert None Field in Word Document
```java
// Get the first section of the document
Section section = document.getSections().get(0);

// Add a new paragraph to the section
Paragraph par = section.addParagraph();

// Append a none field to the paragraph with an empty field code
Field field = par.appendField("", FieldType.Field_None);

// Enable updating fields in the document
document.isUpdateFields(true);
```

---

# Spire.Doc Page Reference Field
## Insert a page reference field into a document
```java
// Get the last section of the document
Section section = document.getLastSection();

// Add a new paragraph to the section
Paragraph par = section.addParagraph();

// Append a page reference field to the paragraph with the specified field name
Field field = par.appendField("pageRef", FieldType.Field_Page_Ref);

// Set the code for the page reference field with specified parameters
field.setCode("PAGEREF bookmark1 \\# \"0\" \\* Arabic \\* MERGEFORMAT");

// Enable updating fields in the document
document.isUpdateFields(true);
```

---

# Spire.Doc Custom Property Fields
## Remove custom property fields from a Word document
```java
// Get the custom document properties of the document
CustomDocumentProperties cdp = document.getCustomDocumentProperties();

// Loop through the custom document properties and remove them
for (int i = 0; i < cdp.getCount(); ) {
    cdp.remove(cdp.get(i).getName());
}

// Enable updating fields in the document
document.isUpdateFields(true);
```

---

# Spire.Doc Field Removal
## Remove a field from a document
```java
// Get the first field in the document
Field field = document.getFields().get(0);

// Get the owner paragraph of the field
Paragraph par = field.getOwnerParagraph();

// Get the index of the field within its parent paragraph
int index = par.getChildObjects().indexOf(field);

// Remove the field from its parent paragraph
par.getChildObjects().removeAt(index);
```

---

# Spire.Doc Text Replacement
## Replace text with merge field
```java
// Find the text "Test" in the document and select it
TextSelection ts = document.findString("Test", true, true);

// Get the selected text range as a single range
TextRange tr = ts.getAsOneRange();

// Get the owner paragraph of the text range
Paragraph par = tr.getOwnerParagraph();

// Get the index of the text range within its parent paragraph
int index = par.getChildObjects().indexOf(tr);

// Create a new merge field
MergeField field = new MergeField(document);
field.setFieldName("MergeField");

// Insert the merge field at the position of the text range within the paragraph
par.getChildObjects().insert(index, field);

// Remove the text range from its parent paragraph
par.getChildObjects().remove(tr);
```

---

# Spire.Doc Field Culture Setting
## Set culture for document fields using locale ID
```java
// Create a regular date field
Field field1 = paragraph.appendField("Date1", FieldType.Field_Date);
field1.setCode("DATE \\@" + "\"yyyy\\MM\\dd\"");

// Create a date field with French culture
Field field2 = newParagraph.appendField("\"\\@\"dd MMMM yyyy", FieldType.Field_Date);
field2.getCharacterFormat().setLocaleIdASCII((short) 1036);

// Enable updating fields in the document
document.isUpdateFields(true);
```

---

# Spire.Doc Set Locale for Field
## Set the locale ID and text for a date field in a document
```java
// Add a new paragraph to the section
Paragraph par = section.addParagraph();

// Append a date field with the specified field name to the paragraph
Field field = par.appendField("DocDate", FieldType.Field_Date);

// Get the text range of the field and set its ASCII locale ID to Russian (1049)
TextRange range = (TextRange) field.getOwnerParagraph().getChildObjects().get(0);
range.getCharacterFormat().setLocaleIdASCII((short) 1049);

// Set the field text to "2019-10-10"
field.setFieldText("2019-10-10");

// Enable updating fields in the document
document.isUpdateFields(true);
```

---

# Document Field Update
## Update fields in a document using Spire.Doc library
```java
// Load an existing document
Document document = new Document();

// Setting the culture source when updating fields
document.getFieldOptions().setCultureSource(FieldCultureSource.CurrentThread);

// Enable updating fields in the document
document.isUpdateFields(true);
```

---

# Change TOC Style in Word Document
## Create and apply custom style to Table of Contents
```java
// Create a custom Table of Contents (TOC) style and set its properties
ParagraphStyle tocStyle = (ParagraphStyle) Style.createBuiltinStyle(BuiltinStyle.Toc_1, doc);
tocStyle.getCharacterFormat().setFontName("Aleo");
tocStyle.getCharacterFormat().setFontSize(15f);
tocStyle.getCharacterFormat().setTextColor(Color.LIGHT_GRAY);

// Add the custom TOC style to the document's styles collection
doc.getStyles().add(tocStyle);

// Iterate through the sections in the document
for (int s = 0; s < doc.getSections().getCount(); s++) {
    Section section = doc.getSections().get(s);

    // Iterate through the child objects in the section's body
    for (int i = 0; i < section.getBody().getChildObjects().getCount(); i++) {
        DocumentObject obj = section.getBody().getChildObjects().get(i);

        // Check if the object is a StructureDocumentTag (SDT)
        if (obj instanceof StructureDocumentTag) {
            StructureDocumentTag tag = (StructureDocumentTag) obj;

            // Iterate through the child objects in the SDT
            for (int j = 0; j < tag.getChildObjects().getCount(); j++) {
                DocumentObject cObj = tag.getChildObjects().get(j);

                // Check if the child object is a paragraph
                if (cObj instanceof Paragraph) {
                    Paragraph para = (Paragraph) cObj;

                    // Check if the paragraph has the style name "TOC1"
                    if (para.getStyleName().equals("TOC1")) {
                        // Apply the custom TOC style to the paragraph
                        para.applyStyle(tocStyle.getName());
                    }
                }
            }
        }
    }
}
```

---

# Change TOC Tab Style
## Modify table of contents tab positions and leaders in Word document
```java
// Iterate through sections in the document
for (int s = 0; s < doc.getSections().getCount(); s++) {
    // Get the current section
    Section section = doc.getSections().get(s);

    // Iterate through child objects in the section's body
    for (int i = 0; i < section.getBody().getChildObjects().getCount(); i++) {
        // Get the current child object
        DocumentObject obj = section.getBody().getChildObjects().get(i);

        // Check if the child object is a StructureDocumentTag
        if (obj instanceof StructureDocumentTag) {
            // Cast the child object to StructureDocumentTag
            StructureDocumentTag tag = (StructureDocumentTag) obj;

            // Iterate through child objects in the StructureDocumentTag
            for (int j = 0; j < tag.getChildObjects().getCount(); j++) {
                // Get the current child object
                DocumentObject cObj = tag.getChildObjects().get(j);

                // Check if the child object is a Paragraph
                if (cObj instanceof Paragraph) {
                    // Cast the child object to Paragraph
                    Paragraph para = (Paragraph) cObj;

                    // Check if the Paragraph has a specific style name
                    if (para.getStyleName().equals("TOC2")) {
                        // Iterate through tabs in the Paragraph's format
                        for (int t = 0; t < para.getFormat().getTabs().getCount(); t++) {
                            // Get the current tab
                            Tab tab = para.getFormat().getTabs().get(t);

                            // Adjust the position of the tab
                            tab.setPosition(tab.getPosition() + 20);

                            // Set the tab leader to No Leader
                            tab.setTabLeader(TabLeader.No_Leader);
                        }
                    }
                }
            }
        }
    }
}
```

---

# Spire.Doc Table of Contents Creation
## Create a table of contents with default settings in a Word document
```java
// Create a new Document object
Document doc = new Document();

// Add a section to the document
Section section = doc.addSection();

// Add a paragraph to the section
Paragraph para = section.addParagraph();

// Append a table of contents to the paragraph
para.appendTOC(1, 3);

// Add another paragraph to the section
Paragraph par = section.addParagraph();

// Set the horizontal alignment of the paragraph to center
par.getFormat().setHorizontalAlignment(HorizontalAlignment.Center);

// Add another paragraph to the section
Paragraph para1 = section.addParagraph();

// Apply the built-in style "Heading 1" to the paragraph
para1.applyStyle(BuiltinStyle.Heading_1);

// Add another paragraph to the section
Paragraph para2 = section.addParagraph();

// Apply the built-in style "Heading 2" to the paragraph
para2.applyStyle(BuiltinStyle.Heading_2);

// Add another paragraph to the section
Paragraph para3 = section.addParagraph();

// Apply the built-in style "Heading 3" to the paragraph
para3.applyStyle(BuiltinStyle.Heading_3);

// Update the table of contents in the document
doc.updateTableOfContents();
```

---

# Spire.Doc Table of Contents Customization
## Create and customize a table of contents in a Word document
```java
// Create a new Document object
Document doc = new Document();

// Add a section to the document
Section section = doc.addSection();

// Create a TableOfContent object with specified format and add it to a paragraph in the section
TableOfContent toc = new TableOfContent(doc, "{\\o \"1-3\" \\n 1-1}");
Paragraph para = section.addParagraph();
para.getItems().add(toc);

// Add field marks to define the start and end of the table of contents
para.appendFieldMark(FieldMarkType.Field_Separator);
para.appendText("TOC");
para.appendFieldMark(FieldMarkType.Field_End);

// Set the document's table of contents to the created TableOfContent object
doc.setTOC(toc);

// Add a paragraph with heading style that will appear in the TOC
Paragraph para1 = section.addParagraph();
para1.appendText("Ornithogalum");
para1.applyStyle(BuiltinStyle.Heading_1);

// Update the table of contents in the document
doc.updateTableOfContents();
```

---

# Spire.Doc Table of Contents Removal
## Remove table of content paragraphs from a Word document
```java
// Get the body of the first section in the document
Body body = document.getSections().get(0).getBody();

// Define a pattern to match the style names of table of contents paragraphs
String pattern = "TOC\\w+";

// Iterate over paragraphs in the body
for (int i = 0; i < body.getParagraphs().getCount(); i++) {
    // Get the style name of the current paragraph
    String styleName = body.getParagraphs().get(i).getStyleName();
    
    // Check if the style name matches the defined pattern
    if (Pattern.matches(pattern, styleName)) {
        // If it matches, remove the paragraph from the body
        body.getParagraphs().removeAt(i);
        i--;
    }
}
```

---

# spire.doc textbox table manipulation
## delete table from textbox
```java
// Get the first textbox
TextBox textbox = doc.getTextBoxes().get(0);

// Remove the first table from the textbox
textbox.getBody().getTables().removeAt(0);
```

---

# Spire.Doc Text Box Text Extraction
## Extract text from text boxes in a Word document
```java
// Create and load the document
Document document = new Document();
document.loadFromFile("data/ExtractTextFromTextBoxes.docx");

// Check if the document has text boxes
if (document.getTextBoxes().getCount() > 0) {
    // Iterate through all sections
    for (int s = 0; s < document.getSections().getCount(); s++) {
        Section section = document.getSections().get(s);

        // Iterate through all paragraphs in the section
        for (int i = 0; i < section.getParagraphs().getCount(); i++) {
            Paragraph p = section.getParagraphs().get(i);

            // Iterate through all child objects in the paragraph
            for (int j = 0; j < p.getChildObjects().getCount(); j++) {
                DocumentObject obj = p.getChildObjects().get(j);

                // Check if the object is a text box
                if (obj.getDocumentObjectType() == DocumentObjectType.Text_Box) {
                    TextBox textbox = (TextBox) obj;

                    // Iterate through all child objects in the text box
                    for (int k = 0; k < textbox.getChildObjects().getCount(); k++) {
                        DocumentObject objt = textbox.getChildObjects().get(k);

                        // Extract text from paragraphs
                        if (objt.getDocumentObjectType() == DocumentObjectType.Paragraph) {
                            String text = ((Paragraph) objt).getText();
                        }

                        // Extract text from tables
                        if (objt.getDocumentObjectType() == DocumentObjectType.Table) {
                            Table table = (Table) objt;
                            extractTextFromTables(table);
                        }
                    }
                }
            }
        }
    }
}

// Dispose the document
document.dispose();

// Helper method to extract text from tables
static void extractTextFromTables(Table table) {
    // Iterate through all rows
    for (int i = 0; i < table.getRows().getCount(); i++) {
        TableRow row = table.getRows().get(i);

        // Iterate through all cells in the row
        for (int j = 0; j < row.getCells().getCount(); j++) {
            TableCell cell = row.getCells().get(j);

            // Iterate through all paragraphs in the cell
            for (int k = 0; k < cell.getParagraphs().getCount(); k++) {
                Paragraph paragraph = cell.getParagraphs().get(k);
                String text = paragraph.getText();
            }
        }
    }
}
```

---

# Spire.Doc Image in TextBox
## Insert an image into a text box in a Word document
```java
// Create a new text box within the paragraph
TextBox tb = paragraph.appendTextBox(220, 220);

// Set the horizontal and vertical position of the text box to the center of the page
tb.getFormat().setHorizontalOrigin(HorizontalOrigin.Page);
tb.getFormat().setHorizontalPosition(50);
tb.getFormat().setVerticalOrigin(VerticalOrigin.Page);
tb.getFormat().setVerticalPosition(50);

//Set the fill effect of textbox as picture
tb.getFormat().getFillEffects().setType(BackgroundType.Picture);

//Fill the textbox with a picture
tb.getFormat().getFillEffects().setPicture("data/spire.Doc.png");
```

---

# Spire.Doc Table in Text Box
## Insert a table into a text box in a Word document
```java
// Create a new Document object
Document doc = new Document();

// Add a new section to the document
Section section = doc.addSection();

// Add a paragraph to the section
Paragraph paragraph = section.addParagraph();

// Append a text box with width 300 and height 100 to the paragraph
TextBox textbox = paragraph.appendTextBox(300, 100);

// Set the horizontal origin and position of the text box relative to the page
textbox.getFormat().setHorizontalOrigin(HorizontalOrigin.Page);
textbox.getFormat().setHorizontalPosition(140);

// Set the vertical origin and position of the text box relative to the page
textbox.getFormat().setVerticalOrigin(VerticalOrigin.Page);
textbox.getFormat().setVerticalPosition(50);

// Add a paragraph to the text box body and set the text to "Table 1"
Paragraph textboxParagraph = textbox.getBody().addParagraph();
TextRange textboxRange = textboxParagraph.appendText("Table 1");

// Set the font name of the text to Arial
textboxRange.getCharacterFormat().setFontName("Arial");

// Add a table to the text box body
Table table = textbox.getBody().addTable(true);

// Reset the table to have 4 rows and 4 columns
table.resetCells(4, 4);

// Loop through the table data and add it to the table cells
for (int i = 0; i < 4; i++) {
    for (int j = 0; j < 4; j++) {
        TextRange tableRange = table.getRows().get(i).getCells().get(j).addParagraph().appendText("Data");
        
        // Set the font name of the text in the table cells to Arial
        tableRange.getCharacterFormat().setFontName("Arial");
    }
}

// Apply a style called "Table_Colorful_2" to the table
table.applyStyle(DefaultTableStyle.Table_Colorful_2);
```

---

# Read Table from TextBox in Word Document
## Extract table data from a TextBox in a Word document and write to a text file
```java
// Create a new Document object
Document doc = new Document();

// Load a Word document into the Document object
doc.loadFromFile("data/textBoxTable.docx");

// Get the first TextBox object from the Document
TextBox textbox = doc.getTextBoxes().get(0);

// Get the first Table object from the TextBox
Table table = textbox.getBody().getTables().get(0);

// Loop through each row in the table
for (int i = 0; i < table.getRows().getCount(); i++) {
    // Get the current row
    TableRow row = table.getRows().get(i);

    // Loop through each cell in the row
    for (int j = 0; j < row.getCells().getCount(); j++) {
        // Get the current cell
        TableCell cell = row.getCells().get(j);

        // Loop through each paragraph in the cell
        for (int k = 0; k < cell.getParagraphs().getCount(); k++) {
            // Get the current paragraph
            Paragraph paragraph = cell.getParagraphs().get(k);

            // Write the text of the paragraph to the file, separated by a tab character
            bw.write(paragraph.getText() + "\t");
        }
    }

    // Write a new line character to the file after each row
    bw.write("\r\n");
}
```

---

# Spire.Doc textbox removal
## Remove a textbox from a Word document
```java
//Remove the first text box
doc.getTextBoxes().removeAt(0);
```

---

# Spire.Doc Textbox Creation
## Create and format textboxes in Word documents
```java
// Insert TextBox with specific formatting
TextBox textBox = paragraph.appendTextBox(240, 35);
textBox.getFormat().setHorizontalAlignment(ShapeHorizontalAlignment.Left);
textBox.getFormat().setLineColor(Color.GRAY);
textBox.getFormat().setLineStyle(TextBoxLineStyle.Simple);
textBox.getFormat().setFillColor(Color.GREEN);

// Add a paragraph inside TextBox and set its text content and formatting
Paragraph para = textBox.getBody().addParagraph();
TextRange txtrg = para.appendText("Textbox in the document");
txtrg.getCharacterFormat().setFontName("Lucida Sans Unicode");
txtrg.getCharacterFormat().setFontSize(14);
txtrg.getCharacterFormat().setTextColor(Color.white);
para.getFormat().setHorizontalAlignment(HorizontalAlignment.Center);
```

---

# Spire.Doc Text Box Formatting
## Create and format a text box with position, border, and margins
```java
// Create a new TextBox object and append it to the Section
TextBox TB = sec.addParagraph().appendTextBox(310, 90);

// Add a new Paragraph to the TextBox
Paragraph para = TB.getBody().addParagraph();

// Append some text to the Paragraph
TextRange TR = para.appendText("Using Spire.Doc, developers will find " + "a simple and effective method to endow their applications with rich MS Word features. ");

// Set the font name and size for the text
TR.getCharacterFormat().setFontName("Cambria ");
TR.getCharacterFormat().setFontSize(13);

// Set the position and size of the TextBox on the page
TB.getFormat().setHorizontalOrigin(HorizontalOrigin.Page);
TB.getFormat().setHorizontalPosition(120);
TB.getFormat().setVerticalOrigin(VerticalOrigin.Page);
TB.getFormat().setVerticalPosition(100);

// Set the style of the TextBox border
TB.getFormat().setLineStyle(TextBoxLineStyle.Double);
TB.getFormat().setLineColor(Color.BLUE);
TB.getFormat().setLineDashing(LineDashing.Solid);
TB.getFormat().setLineWidth(5);

// Set the internal margins of the TextBox
TB.getFormat().getInternalMargin().setTop(15);
TB.getFormat().getInternalMargin().setBottom(10);
TB.getFormat().getInternalMargin().setLeft(12);
TB.getFormat().getInternalMargin().setRight(10);
```

---

# Spire.Doc Image Watermark
## Add an image watermark to a document
```java
public static void insertImageWatermark(Document document) {
    // Create a new instance of the PictureWatermark class with the input image file name
    PictureWatermark picture = new PictureWatermark();
    picture.setPicture("data/imageWatermark.png");

    // Set the scaling factor for the watermark image
    picture.setScaling(250);

    // Set whether the watermark should be washed out or not
    picture.isWashout(false);

    // Set the watermark to be applied to the document
    document.setWatermark(picture);
}
```

---

# Spire.Doc Picture Watermark
## Add a picture watermark to a document using BufferedImage
```java
// Create a new instance of the PictureWatermark class with the input BufferedImage
PictureWatermark picture = new PictureWatermark(bufferedImage, false);

// Set the scaling factor for the watermark image
picture.setScaling(250);

// Set the watermark to be applied to the document
document.setWatermark(picture);
```

---

# Remove Watermark in Word Document
## This code demonstrates how to remove watermark from a Word document
```java
//Create Word document
Document document = new Document();

//Set the watermark as null to remove the text and image watermark
document.setWatermark(null);
```

---

# Spire.Doc Text Watermark
## Add text watermark to Word document
```java
public static void insertTextWatermark(Section section) {
    // Create a TextWatermark object
    TextWatermark txtWatermark = new TextWatermark();

    // Set the text content of the watermark
    txtWatermark.setText("E-iceblue");

    // Set the font size of the watermark
    txtWatermark.setFontSize(25);

    // Set the color of the watermark
    txtWatermark.setColor(Color.blue);

    // Set the layout of the watermark to Diagonal
    txtWatermark.setLayout(WatermarkLayout.Diagonal);

    // Set the watermark for the document's section
    section.getDocument().setWatermark(txtWatermark);
}
```

---

# Spire.Doc OLE Object Extraction
## Extract OLE objects from Word documents based on their types
```java
// Traverse through all sections of the document
for (int s = 0; s < doc.getSections().getCount(); s++) {
    Section section = doc.getSections().get(s);

    // Traverse through all child objects in the body of each section
    for (int i = 0; i < section.getBody().getChildObjects().getCount(); i++) {
        DocumentObject obj = section.getBody().getChildObjects().get(i);

        // Check whether the object is a paragraph
        if (obj instanceof Paragraph) {
            Paragraph par = (Paragraph) obj;

            // Traverse through all child objects in the paragraph
            for (int j = 0; j < par.getChildObjects().getCount(); j++) {
                DocumentObject o = par.getChildObjects().get(j);

                // Check whether the object is an OLE object
                if (o.getDocumentObjectType() == DocumentObjectType.Ole_Object) {
                    DocOleObject ole = (DocOleObject) o;

                    // Get the type of the OLE object
                    String type = ole.getObjectType();

                    // Check whether the object type is "Acrobat.Document.11"
                    if ("AcroExch.Document.DC".equals(type)) {
                        // Write the data of the OLE object to a PDF file
                        byteArrayToFile(ole.getNativeData(), "output/extractOLE.pdf");
                    }

                    // Check whether the object type is "Excel.Sheet.8"
                    else if ("Excel.Sheet.8".equals(type)) {
                        // Write the data of the OLE object to an Excel file
                        byteArrayToFile(ole.getNativeData(), "output/extractOLE.xls");
                    }

                    // Check whether the object type is "PowerPoint.Show.12"
                    else if ("PowerPoint.Show.12".equals(type)) {
                        // Write the data of the OLE object to a PowerPoint file
                        byteArrayToFile(ole.getNativeData(), "output/extractOLE.pptx");
                    }
                }
            }
        }
    }
}

public static void byteArrayToFile(byte[] datas, String destPath) {
    // Create a File object with the destination path
    File dest = new File(destPath);

    try (
            // Create an InputStream from the byte array
            InputStream is = new ByteArrayInputStream(datas);

            // Create an OutputStream to write data to the file
            OutputStream os = new BufferedOutputStream(new FileOutputStream(dest, false));
    ) {
        // Create a buffer to read data in chunks
        byte[] flush = new byte[1024];
        int len = -1;

        // Read data from the InputStream and write it to the OutputStream
        while ((len = is.read(flush)) != -1) {
            os.write(flush, 0, len);
        }

        // Flush any remaining data in the OutputStream
        os.flush();
    } catch (IOException e) {
        e.printStackTrace();
    }
}
```

---

# Spire.Doc OLE Object Insertion
## Insert OLE object into Word document
```java
// Create a document
Document doc = new Document();

// Add a section
Section sec = doc.addSection();

// Add a paragraph
Paragraph par = sec.addParagraph();

// Create a DocPicture
DocPicture picture = new DocPicture(doc);

// Insert the OLE
DocOleObject obj = par.appendOleObject("data/example.xlsx", picture, OleObjectType.Excel_Worksheet);
```

---

# Spire.Doc CheckBox Content Control
## Add CheckBox Content Control to Word Document
```java
//Create a document
Document document = new Document();

//Add a new section.
Section section = document.addSection();

//Add a paragraph
Paragraph paragraph = section.addParagraph();

//Create StructureDocumentTagInline for document
StructureDocumentTagInline sdt = new StructureDocumentTagInline(document);

//Add sdt in paragraph
paragraph.getChildObjects().add(sdt);

//Specify the type
sdt.getSDTProperties().setSDTType(SdtType.Check_Box);

//Set properties for control
SdtCheckBox scb = new SdtCheckBox();
sdt.getSDTProperties().setControlProperties(scb);

//Add textRange format
TextRange tr = new TextRange(document);
tr.getCharacterFormat().setFontName("MS Gothic");
tr.getCharacterFormat().setFontSize(12);

//Add textRange to StructureDocumentTagInline
sdt.getChildObjects().add(tr);

//Set checkBox as checked
scb.setChecked(true);
```

---

# Adding Content Controls to Word Documents
## This code demonstrates how to add different types of content controls (Combo Box, Text, Picture, Date Picker, and Drop-Down List) to a Word document using Spire.Doc for Java.

```java
// Create a StructureDocumentTagInline object for Combo Box
StructureDocumentTagInline sd = new StructureDocumentTagInline(document);
paragraph.getChildObjects().add(sd);
sd.getSDTProperties().setSDTType(SdtType.Combo_Box);

// Create and configure a SdtComboBox
SdtComboBox cb = new SdtComboBox();
cb.getListItems().add(new SdtListItem("Spire.Doc"));
cb.getListItems().add(new SdtListItem("Spire.XLS"));
cb.getListItems().add(new SdtListItem("Spire.PDF"));
sd.getSDTProperties().setControlProperties(cb);

// Add text to the combo box
TextRange rt = new TextRange(document);
rt.setText(cb.getListItems().get(0).getDisplayText());
sd.getSDTContent().getChildObjects().add(rt);

// Create a StructureDocumentTagInline object for Text
sd = new StructureDocumentTagInline(document);
paragraph.getChildObjects().add(sd);
sd.getSDTProperties().setSDTType(SdtType.Text);

// Create and configure a SdtText
SdtText text = new SdtText(true);
text.isMultiline(true);
sd.getSDTProperties().setControlProperties(text);

// Add text to the text control
rt = new TextRange(document);
rt.setText("Text");
sd.getSDTContent().getChildObjects().add(rt);

// Create a StructureDocumentTagInline object for Picture
sd = new StructureDocumentTagInline(document);
paragraph.getChildObjects().add(sd);
sd.getSDTProperties().setSDTType(SdtType.Picture);

// Create and configure a DocPicture
DocPicture pic = new DocPicture(document);
pic.setWidth(10f);
pic.setHeight(10f);
sd.getSDTContent().getChildObjects().add(pic);

// Create a StructureDocumentTagInline object for Date Picker
sd = new StructureDocumentTagInline(document);
paragraph.getChildObjects().add(sd);
sd.getSDTProperties().setSDTType(SdtType.Date_Picker);

// Create and configure a SdtDate
SdtDate date = new SdtDate();
date.setCalendarType(CalendarType.Default);
date.setDateFormat("yyyy.MM.dd");
date.setFullDate(new Date());
sd.getSDTProperties().setControlProperties(date);

// Add text to the date picker
rt = new TextRange(document);
rt.setText("2019.12.31");
sd.getSDTContent().getChildObjects().add(rt);

// Create a StructureDocumentTagInline object for Drop-Down List
sd = new StructureDocumentTagInline(document);
paragraph.getChildObjects().add(sd);
sd.getSDTProperties().setSDTType(SdtType.Drop_Down_List);

// Create and configure a SdtDropDownList
SdtDropDownList sddl = new SdtDropDownList();
sddl.getListItems().add(new SdtListItem("Harry"));
sddl.getListItems().add(new SdtListItem("Jerry"));
sd.getSDTProperties().setControlProperties(sddl);

// Add text to the drop-down list
rt = new TextRange(document);
rt.setText(sddl.getListItems().get(0).getDisplayText());
sd.getSDTContent().getChildObjects().add(rt);
```

---

# Spire.Doc RichText Content Control
## Add RichText content control to Word document
```java
//Create StructureDocumentTagInline for document
StructureDocumentTagInline sdt = new StructureDocumentTagInline(document);

//Add sdt in paragraph
paragraph.getChildObjects().add(sdt);

//Specify the type
sdt.getSDTProperties().setSDTType(SdtType.Rich_Text);

//Set displaying text
SdtText text = new SdtText(true);

// Enable multiline
text.isMultiline(true);

// Set the control properties of the StructureDocumentTagInline to the SdtText
sdt.getSDTProperties().setControlProperties(text);

//Crate a TextRange
TextRange rt = new TextRange(document);

// Append text
rt.setText("Welcome to use ");

// Set color
rt.getCharacterFormat().setTextColor(Color.GREEN);

// Add the text range
sdt.getSDTContent().getChildObjects().add(rt);

// create a new TextRange
rt = new TextRange(document);

// Append text
rt.setText("Spire.Doc");

// Set color
rt.getCharacterFormat().setTextColor(Color.ORANGE);

// Add the text range
sdt.getSDTContent().getChildObjects().add(rt);
```

---

# Spire.Doc ComboBox Item Manipulation
## Modify combo box items in a Word document by removing, adding, and selecting items
```java
// Iterate through sections in the document
for (int i = 0; i < doc.getSections().getCount(); i++) {
    Section section = doc.getSections().get(i);

    // Iterate through child objects in the body of each section
    for (int j = 0; j < section.getBody().getChildObjects().getCount(); j++) {
        DocumentObject bodyObj = section.getBody().getChildObjects().get(j);

        // Check if the object is a structure document tag
        if (bodyObj.getDocumentObjectType() == DocumentObjectType.Structure_Document_Tag) {
            StructureDocumentTag sdt = (StructureDocumentTag) bodyObj;

            // Check if the structure document tag is a combo box
            if (sdt.getSDTProperties().getSDTType() == SdtType.Combo_Box) {
                SdtComboBox combo = (SdtComboBox) sdt.getSDTProperties().getControlProperties();

                // Remove an item from the combo box list
                combo.getListItems().removeAt(1);

                // Add a new item to the combo box list
                SdtListItem item = new SdtListItem("D", "D");

                // Add the newly created item
                combo.getListItems().add(item);

                // Select the item with value "D"
                for (int k = 0; k < combo.getListItems().getCount(); k++) {
                    SdtListItem sdtItem = combo.getListItems().get(k);
                    if ("D".equals(sdtItem.getValue())) {
                        // Select the item
                        combo.getListItems().setSelectedValue(sdtItem);
                    }
                }
            }
        }
    }
}
```

---

# Spire.Doc Content Control Properties
## Extract properties from structured document tags in a Word document
```java
// Get all structure tags from the document
structureTags structureTags = GetAllTags(doc);

// Declare variables for storing properties
String alias;
BigDecimal id;
String tag;
String sdtType;
SdtType sdt = SdtType.Rich_Text;
String content = "";

// Process inline structure tags
List<StructureDocumentTagInline> tagInlines = structureTags.getM_tagInlines();
for (int i = 0; i < tagInlines.size(); i++) {
    // Retrieve properties of the structure tag
    alias = tagInlines.get(i).getSDTProperties().getAlias();
    id = tagInlines.get(i).getSDTProperties().getId();
    tag = tagInlines.get(i).getSDTProperties().getTag();
    sdt = tagInlines.get(i).getSDTProperties().getSDTType();
    sdtType = sdt.toString();

    // Check if the structure tag contains rich text or plain text
    if (sdt == SdtType.Rich_Text || sdt == SdtType.Text) {
        if (tagInlines.get(i).getChildObjects().getCount() > 0) {
            // Iterate through child objects within the structure tag
            for (int k = 0; k < tagInlines.get(i).getChildObjects().getCount(); k++) {
                if (tagInlines.get(i).getChildObjects().get(k).getDocumentObjectType() == DocumentObjectType.Text_Range) {
                    // Retrieve text content from a text range object
                    TextRange textRange = (TextRange) tagInlines.get(i).getChildObjects().get(k);
                    content += textRange.getText();
                }
            }
        }
    }

    // Reset the content variable
    content = "";
}

// Process other structure tags
List<StructureDocumentTag> tags = structureTags.getM_tags();
for (int i = 0; i < tags.size(); i++) {
    // Retrieve properties of the structure tag
    alias = tags.get(i).getSDTProperties().getAlias();
    id = tags.get(i).getSDTProperties().getId();
    tag = tags.get(i).getSDTProperties().getTag();
    sdt = tags.get(i).getSDTProperties().getSDTType();
    sdtType = sdt.toString();

    // Check if the structure tag contains rich text or plain text
    if (sdt == SdtType.Rich_Text || sdt == SdtType.Text) {
        if (tags.get(i).getChildObjects().getCount() > 0) {
            // Iterate through child objects within the structure tag
            for (int k = 0; k < tags.get(i).getChildObjects().getCount(); k++) {
                if (tags.get(i).getChildObjects().get(k).getDocumentObjectType() == DocumentObjectType.Paragraph) {
                    // Retrieve text content from a paragraph object
                    Paragraph paragraph = (Paragraph) tags.get(i).getChildObjects().get(k);
                    content += paragraph.getText();
                }
            }
        }
    }

    // Reset the content variable
    content = "";
}

// Get all the row tags from structureTags
List<StructureDocumentTagRow> rowtags = structureTags.getM_rowtags();

// Iterate over each row tag
for (int i = 0; i < rowtags.size(); i++) {
    alias = rowtags.get(i).getSDTProperties().getAlias();
    id = rowtags.get(i).getSDTProperties().getId();
    tag = rowtags.get(i).getSDTProperties().getTag();
    sdt = rowtags.get(i).getSDTProperties().getSDTType();
    sdtType = sdt.toString();

    // Check if the SDT type is Rich_Text or Text
    if (sdt == SdtType.Rich_Text || sdt == SdtType.Text) {
        // Check if the row tag has child objects
        if (rowtags.get(i).getChildObjects().getCount() > 0) {
            // Iterate over each child object of the row tag
            for (int k = 0; k < rowtags.get(i).getChildObjects().getCount(); k++) {
                // Check if the child object is a Paragraph
                if (rowtags.get(i).getChildObjects().get(k).getDocumentObjectType() == DocumentObjectType.Paragraph) {
                    Paragraph paragraph = (Paragraph) rowtags.get(i).getChildObjects().get(k);
                    content += paragraph.getText();
                }
            }
        }
    }

    // Reset the content variable
    content = "";
}

// Get all the cell tags from structureTags
List<StructureDocumentTagCell> celltags = structureTags.getM_celltags();

// Iterate over each cell tag
for (int i = 0; i < celltags.size(); i++) {
    alias = celltags.get(i).getSDTProperties().getAlias();
    id = celltags.get(i).getSDTProperties().getId();
    tag = celltags.get(i).getSDTProperties().getTag();
    sdt = celltags.get(i).getSDTProperties().getSDTType();
    sdtType = sdt.toString();

    // Check if the SDT type is Rich_Text or Text
    if (sdt == SdtType.Rich_Text || sdt == SdtType.Text) {
        // Check if the cell tag has child objects
        if (celltags.get(i).getChildObjects().getCount() > 0) {
            // Iterate over each child object of the cell tag
            for (int k = 0; k < celltags.get(i).getChildObjects().getCount(); k++) {
                // Check if the child object is a Paragraph
                if (celltags.get(i).getChildObjects().get(k).getDocumentObjectType() == DocumentObjectType.Paragraph) {
                    Paragraph paragraph = (Paragraph) celltags.get(i).getChildObjects().get(k);
                    content += paragraph.getText();
                }
            }
        }
    }

    // Reset the content variable
    content = "";
}

public static structureTags GetAllTags(Document document) {
    // Create a new instance of StructureTags
    structureTags structureTags = new structureTags();

    // Iterate over each section in the document
    for (int i = 0; i < document.getSections().getCount(); i++) {
        Section section = document.getSections().get(i);

        // Iterate over each child object in the section's body
        for (int j = 0; j < section.getBody().getChildObjects().getCount(); j++) {
            DocumentObject obj = section.getBody().getChildObjects().get(j);

            // Check if the child object is a Structure_Document_Tag
            if (obj.getDocumentObjectType() == DocumentObjectType.Structure_Document_Tag) {
                structureTags.getM_tags().add((StructureDocumentTag) obj);
            }
            // Check if the child object is a Paragraph
            else if (obj.getDocumentObjectType() == DocumentObjectType.Paragraph) {
                Paragraph para = (Paragraph) obj;

                // Iterate over each child object in the paragraph
                for (int k = 0; k < para.getChildObjects().getCount(); k++) {
                    DocumentObject pobj = para.getChildObjects().get(k);

                    // Check if the child object is a Structure_Document_Tag_Inline
                    if (pobj.getDocumentObjectType() == DocumentObjectType.Structure_Document_Tag_Inline) {
                        structureTags.getM_tagInlines().add((StructureDocumentTagInline) pobj);
                    }
                }
            }
            // Check if the child object is a Table
            else if (obj.getDocumentObjectType() == DocumentObjectType.Table) {
                Table table = (Table) obj;

                // Iterate over each row in the table
                for (int r = 0; r < table.getRows().getCount(); r++) {
                    // Check if the row is a Structure_Document_Tag_Row
                    if (table.getRows().get(r).getDocumentObjectType() == DocumentObjectType.Structure_Document_Tag_Row) {
                        structureTags.getM_rowtags().add((StructureDocumentTagRow) (table.getRows().get(r)));
                    }

                    TableRow row = table.getRows().get(r);

                    // Iterate over each cell in the row
                    for (int c = 0; c < row.getCells().getCount(); c++) {
                        // Check if the cell is a Structure_Document_Tag_Cell
                        if (row.getCells().get(c).getDocumentObjectType() == DocumentObjectType.Structure_Document_Tag_Cell) {
                            structureTags.getM_celltags().add((StructureDocumentTagCell) (row.getCells().get(c)));
                        }

                        TableCell cell = row.getCells().get(c);

                        // Iterate over each child object in the cell
                        for (int s = 0; s < cell.getChildObjects().getCount(); s++) {
                            DocumentObject cellChild = cell.getChildObjects().get(s);

                            // Check if the child object is a Structure_Document_Tag
                            if (cellChild.getDocumentObjectType() == DocumentObjectType.Structure_Document_Tag) {
                                structureTags.getM_tags().add((StructureDocumentTag) cellChild);
                            }
                            // Check if the child object is a Paragraph
                            else if (cellChild.getDocumentObjectType() == DocumentObjectType.Paragraph) {
                                Paragraph para = (Paragraph) cellChild;

                                // Iterate over each child object in the paragraph
                                for (int t = 0; t < para.getChildObjects().getCount(); t++) {
                                    DocumentObject pobj = para.getChildObjects().get(t);

                                    // Check if the child object is a Structure_Document_Tag_Inline
                                    if (pobj.getDocumentObjectType() == DocumentObjectType.Structure_Document_Tag_Inline) {
                                        structureTags.getM_tagInlines().add((StructureDocumentTagInline) pobj);
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }
    }

    return structureTags;
}
```

---

# Modify SDT Color in Word Document
## Change structured document tag colors based on their type
```java
// Iterate over each section in the document
for (int s = 0; s < doc.getSections().getCount(); s++) {
    Section section = doc.getSections().get(s);

    // Iterate over each child object in the section's body
    for (int i = 0; i < section.getBody().getChildObjects().getCount(); i++) {

        // Check if the child object is a Paragraph
        if (section.getBody().getChildObjects().get(i) instanceof Paragraph) {
            Paragraph para = (Paragraph) section.getBody().getChildObjects().get(i);

            // Iterate over each child object in the paragraph
            for (int j = 0; j < para.getChildObjects().getCount(); j++) {

                // Check if the child object is a StructureDocumentTagInline
                if (para.getChildObjects().get(j) instanceof StructureDocumentTagInline) {
                    StructureDocumentTagInline sdt = (StructureDocumentTagInline) para.getChildObjects().get(j);
                    SDTProperties sdtProperties = sdt.getSDTProperties();

                    // Update the color based on the SDT type
                    switch (sdtProperties.getSDTType()){
                        case Rich_Text:
                            sdtProperties.setColor(ORANGE);
                            break;
                        case Text:
                            sdtProperties.setColor(GREEN);
                            break;
                    }
                }
            }
        }

        // Check if the child object is a StructureDocumentTag
        if (section.getBody().getChildObjects().get(i) instanceof StructureDocumentTag) {
            StructureDocumentTag sdt = (StructureDocumentTag) section.getBody().getChildObjects().get(i);
            SDTProperties sdtProperties = sdt.getSDTProperties();

            // Update the color based on the SDT type
            switch (sdtProperties.getSDTType()){
                case Rich_Text:
                    sdtProperties.setColor(ORANGE);
                    break;
                case Text:
                    sdtProperties.setColor(GREEN);
                    break;
            }
        }
    }
}
```

---

# Remove Content Controls from Word Document
## This code demonstrates how to remove structured document tags (content controls) from a Word document
```java
// Iterate over each section in the document
for (int s = 0; s < doc.getSections().getCount(); s++) {
    Section section = doc.getSections().get(s);

    // Iterate over each child object in the section's body
    for (int i = 0; i < section.getBody().getChildObjects().getCount(); i++) {

        // Check if the child object is a Paragraph
        if (section.getBody().getChildObjects().get(i) instanceof Paragraph) {
            Paragraph para = (Paragraph) section.getBody().getChildObjects().get(i);

            // Iterate over each child object in the paragraph
            for (int j = 0; j < para.getChildObjects().getCount(); j++) {

                // Check if the child object is a StructureDocumentTagInline
                if (para.getChildObjects().get(j) instanceof StructureDocumentTagInline) {
                    StructureDocumentTagInline sdt = (StructureDocumentTagInline) para.getChildObjects().get(j);

                    // Remove the StructureDocumentTagInline from the paragraph
                    para.getChildObjects().remove(sdt);

                    // Decrement the index to account for the removed object
                    j--;
                }
            }
        }

        // Check if the child object is a StructureDocumentTag
        if (section.getBody().getChildObjects().get(i) instanceof StructureDocumentTag) {
            StructureDocumentTag sdt = (StructureDocumentTag) section.getBody().getChildObjects().get(i);

            // Remove the StructureDocumentTag from the section's body
            section.getBody().getChildObjects().remove(sdt);

            // Decrement the index to account for the removed object
            i--;
        }
    }
}
```

---

# Remove Content Control Tags After Editing
## This code demonstrates how to remove content control tags after editing by setting them as temporary

```java
// Get all the structure tags in the document
structureTags structureTags = GetAllTags(doc);

// Remove ContentControl tags after editing for tagInlines
List<StructureDocumentTagInline> tagInlines = structureTags.getM_tagInlines();
for (int i = 0; i < tagInlines.size(); i++) {
    StructureDocumentTagInline std = tagInlines.get(i);
    std.getSDTProperties().isTemporary(true);
}

// Remove ContentControl tags after editing for tags
List<StructureDocumentTag> tags = structureTags.getM_tags();
for (int i = 0; i < tags.size(); i++) {
    StructureDocumentTag std = tags.get(i);
    std.getSDTProperties().isTemporary(true);
}

// Remove ContentControl tags after editing for rowTags
List<StructureDocumentTagRow> rowTags = structureTags.getM_rowtags();
for (int i = 0; i < rowTags.size(); i++) {
    StructureDocumentTagRow std = rowTags.get(i);
    std.getSDTProperties().isTemporary(true);
}

// Remove ContentControl tags after editing for cellTags
List<StructureDocumentTagCell> cellTags = structureTags.getM_celltags();
for (int i = 0; i < cellTags.size(); i++) {
    StructureDocumentTagCell std = cellTags.get(i);
    std.getSDTProperties().isTemporary(true);
}

public static structureTags GetAllTags(Document document) {
    // Create a StructureTags
    structureTags structureTags = new structureTags();

    // Iterate through the sections of the document
    for (int i = 0; i < document.getSections().getCount(); i++) {
        Section section = document.getSections().get(i);

        // Iterate through the child objects in the body of each section
        for (int j = 0; j < section.getBody().getChildObjects().getCount(); j++) {
            DocumentObject obj = section.getBody().getChildObjects().get(j);

            // Check the type of the child object
            if (obj.getDocumentObjectType() == DocumentObjectType.Structure_Document_Tag) {
                // If it is a StructureDocumentTag, add it to the list of tags
                structureTags.getM_tags().add((StructureDocumentTag) obj);
            } else if (obj.getDocumentObjectType() == DocumentObjectType.Paragraph) {
                // If it is a Paragraph, iterate through its child objects
                Paragraph para = (Paragraph) obj;
                for (int k = 0; k < para.getChildObjects().getCount(); k++) {
                    DocumentObject pobj = para.getChildObjects().get(k);
                    if (pobj.getDocumentObjectType() == DocumentObjectType.Structure_Document_Tag_Inline) {
                        // If it is a StructureDocumentTagInline, add it to the list of tagInlines
                        structureTags.getM_tagInlines().add((StructureDocumentTagInline) pobj);
                    }
                }
            } else if (obj.getDocumentObjectType() == DocumentObjectType.Table) {
                // If it is a Table, iterate through its rows and cells
                Table table = (Table) obj;
                for (int r = 0; r < table.getRows().getCount(); r++) {
                    if (table.getRows().get(r).getDocumentObjectType() == DocumentObjectType.Structure_Document_Tag_Row) {
                        // If it is a StructureDocumentTagRow, add it to the list of rowTags
                        structureTags.getM_rowtags().add((StructureDocumentTagRow) (table.getRows().get(r)));
                    }
                    TableRow row = table.getRows().get(r);
                    for (int c = 0; c < row.getCells().getCount(); c++) {
                        if (row.getCells().get(c).getDocumentObjectType() == DocumentObjectType.Structure_Document_Tag_Cell) {
                            // If it is a StructureDocumentTagCell, add it to the list of cellTags
                            structureTags.getM_celltags().add((StructureDocumentTagCell) (row.getCells().get(c)));
                        }
                        TableCell cell = row.getCells().get(c);
                        for (int s = 0; s < cell.getChildObjects().getCount(); s++) {
                            DocumentObject cellChild = cell.getChildObjects().get(s);
                            if (cellChild.getDocumentObjectType() == DocumentObjectType.Structure_Document_Tag) {
                                // If it is a StructureDocumentTag, add it to the list of tags
                                structureTags.getM_tags().add((StructureDocumentTag) cellChild);
                            } else if (cellChild.getDocumentObjectType() == DocumentObjectType.Paragraph) {
                                // If it is a Paragraph, iterate through its child objects
                                Paragraph para = (Paragraph) cellChild;
                                for (int t = 0; t < para.getChildObjects().getCount(); t++) {
                                    DocumentObject pobj = para.getChildObjects().get(t);
                                    if (pobj.getDocumentObjectType() == DocumentObjectType.Structure_Document_Tag_Inline) {
                                        // If it is a StructureDocumentTagInline, add it to the list of tagInlines
                                        structureTags.getM_tagInlines().add((StructureDocumentTagInline) pobj);
                                    }
                                }
                            }
                        }
                    }
                }
            }
        }
    }

    // Return the StructureTags object containing all the retrieved tags
    return structureTags;
}

class structureTags {
    private List<StructureDocumentTagInline> m_tagInlines;

    public void setM_tagInlines(List<StructureDocumentTagInline> m_tagInlines) {
        this.m_tagInlines = m_tagInlines;
    }

    public List<StructureDocumentTagInline> getM_tagInlines() {
        if (m_tagInlines == null)
            m_tagInlines = new ArrayList<StructureDocumentTagInline>();
        return m_tagInlines;
    }

    private List<StructureDocumentTag> m_tags;

    public List<StructureDocumentTag> getM_tags() {
        if (m_tags == null)
            m_tags = new ArrayList<StructureDocumentTag>();
        return m_tags;
    }

    public void setM_tags(List<StructureDocumentTag> m_tags) {
        this.m_tags = m_tags;
    }

    private List<StructureDocumentTagRow> m_rowtags;

    public List<StructureDocumentTagRow> getM_rowtags() {
        if (m_rowtags == null)
            m_rowtags = new ArrayList<StructureDocumentTagRow>();
        return m_rowtags;
    }

    public void setM_rowtags(List<StructureDocumentTagRow> m_rowtags) {
        this.m_rowtags = m_rowtags;
    }
    private List<StructureDocumentTagCell> m_celltags;

    public List<StructureDocumentTagCell> getM_celltags() {
        if (m_celltags == null)
            m_celltags = new ArrayList<StructureDocumentTagCell>();
        return m_celltags;
    }

    public void setM_celltags(List<StructureDocumentTagCell> m_celltags) {
        this.m_celltags = m_celltags;
    }
}
```

---

# Spire.Doc content control appearance
## Set appearance of structured document tags in Word document
```java
//Traverse all content controls
for (Object sectionObj : doc.getSections()) {
    Section section = (Section) sectionObj;
    for (Object docObj : section.getBody().getChildObjects()) {
        // Get structureTag in the Word document
        if (docObj instanceof StructureDocumentTag) {
            DocumentObject Obj = (DocumentObject) docObj;
            SDTProperties sdtProperties = ((StructureDocumentTag) Obj).getSDTProperties();
            // Set appearance of the content control
            switch (sdtProperties.getSDTType()) {
                //Set appearance as "Hidden"
                case Text:
                    sdtProperties.setAppearance(SdtAppearance.Hidden);
                    break;
                //Set appearance as "BoundingBox"
                case Rich_Text:
                    sdtProperties.setAppearance(SdtAppearance.Bounding_Box);
                    break;
                //Set appearance as "Tags"
                case Picture:
                    sdtProperties.setAppearance(SdtAppearance.Tags);
                    break;
                //Set appearance as "Default"
                case Check_Box:
                    sdtProperties.setAppearance(SdtAppearance.Default);
                    break;
            }
        }
    }
}
```

---

# Update Checkboxes in Document
## Toggle the checked state of checkbox controls in a structured document
```java
//Call StructureTags
StructureTagInLines structureTags = GetAllTags(document);

//Create list
List<StructureDocumentTagInline> tagInlines = structureTags.getM_tagInlines();

//Get the controls
for (int i = 0; i < tagInlines.size(); i++)
{
    //Get the type
    String type = tagInlines.get(i).getSDTProperties().getSDTType().toString();

    //Update the status
    if ("Check_Box".equals(type))
    {
        // Get the sdtCheckBox
        SdtCheckBox scb = (SdtCheckBox)tagInlines.get(i).getSDTProperties().getControlProperties() ;

        // Judge the status
        if (scb.getChecked())
        {
            scb.setChecked(false);
        }
        else
        {
            scb.setChecked(true);
        }
    }
}

public static StructureTagInLines GetAllTags(Document document) {
    StructureTagInLines structureTags = new StructureTagInLines();

    // Iterate through the sections of the document
    for (int i = 0; i < document.getSections().getCount(); i++) {
        Section section = document.getSections().get(i);

        // Iterate through the child objects in the body of each section
        for (int j = 0; j < section.getBody().getChildObjects().getCount(); j++){
            DocumentObject obj = section.getBody().getChildObjects().get(j);

            // Check if the child object is a paragraph
            if(obj.getDocumentObjectType() == DocumentObjectType.Paragraph) {

                // Iterate through the child objects of the paragraph
                for (int k = 0; k < ((Paragraph)obj).getChildObjects().getCount(); k++) {
                    DocumentObject pobj = ((Paragraph)obj).getChildObjects().get(k);

                    // Check if the child object is a StructureDocumentTagInline
                    if (pobj.getDocumentObjectType() == DocumentObjectType.Structure_Document_Tag_Inline) {
                        // If it is, add it to the list of tagInlines
                        structureTags.getM_tagInlines().add((StructureDocumentTagInline)pobj);
                    }
                }
            }
        }
    }

    // Return the StructureTagInLines object containing all the retrieved StructureDocumentTagInline objects
    return structureTags;
}

class StructureTagInLines {
    private List<StructureDocumentTagInline> m_tagInlines;

    public void setM_tagInlines(List<StructureDocumentTagInline> m_tagInlines) {
        this.m_tagInlines = m_tagInlines;
    }

    public List<StructureDocumentTagInline> getM_tagInlines() {
        if (m_tagInlines == null)
            m_tagInlines = new ArrayList<StructureDocumentTagInline>();
        return m_tagInlines;
    }
}
```

---

# Adding Math Equations to Document
## Demonstrates how to add mathematical equations in LaTeX and MathML formats to a Word document
```java
// Create a document
Document doc = new Document();

// Get the first section
Section section = doc.getSections().get(0);

Paragraph paragraph = null;
OfficeMath officeMath;

// Get the first table and add LaTeX code
Table table1 = section.getTables().get(0);
for (int i = 0; i < latexMathCode.length; i++) {
    paragraph = table1.getRows().get(i + 1).getCells().get(0).addParagraph();
    paragraph.setText(latexMathCode[i]);
    paragraph = table1.getRows().get(i + 1).getCells().get(1).addParagraph();
    officeMath = new OfficeMath(doc);
    officeMath.fromLatexMathCode(latexMathCode[i]);
    paragraph.getItems().add(officeMath);
}

// Get the second table and add MathML code
Table table2 = section.getTables().get(1);
for (int i = 0; i < mathMLCode.length; i++) {
    paragraph = table2.getRows().get(i + 1).getCells().get(0).addParagraph();
    paragraph.setText(mathMLCode[i]);
    paragraph = table2.getRows().get(i + 1).getCells().get(1).addParagraph();
    officeMath = new OfficeMath(doc);
    officeMath.fromMathMLCode(mathMLCode[i]);
    paragraph.getItems().add(officeMath);
}
```

---

# Extract Math Equations from Word Document
## Iterate through document sections, paragraphs and child objects to find OfficeMath equations and convert them to MathML code
```java
// Create a StringBuilder to store the MathML code
StringBuilder stringBuilder = new StringBuilder();

// Iterate over sections in the document
for (int i = 0; i < doc.getSections().getCount(); i++) {
    // Iterate over paragraphs in each section
    for (int j = 0; j < doc.getSections().get(i).getParagraphs().getCount(); j++) {
        // Iterate over child objects in each paragraph
        for (int k = 0; k < doc.getSections().get(i).getParagraphs().get(j).getChildObjects().getCount(); k++) {
            // Get the current DocumentObject
            DocumentObject obj = doc.getSections().get(i).getParagraphs().get(j).getChildObjects().get(k);

            // Check if the object is an OfficeMath equation
            if (obj instanceof OfficeMath) {
                // Cast the object to OfficeMath
                OfficeMath math = (OfficeMath) obj;

                // Append the MathML code of the equation to the StringBuilder
                stringBuilder.append(math.toMathMLCode()).append("\n");
            }
        }
    }
}
```

---

# Spire.Doc Endnote Insertion
## Insert and format endnotes in a Word document
```java
// Create a Document object
Document doc = new Document();

// Get the first section from the document
Section s = doc.getSections().get(0);

// Get the second paragraph from the section
Paragraph p = s.getParagraphs().get(1);

// Add an endnote to the paragraph
Footnote endnote = p.appendFootnote(FootnoteType.Endnote);

// Append text to the endnote's text body
TextRange text = endnote.getTextBody().addParagraph().appendText("Reference: Wikipedia");

// Set the format of the text in the endnote
text.getCharacterFormat().setFontName("Impact");
text.getCharacterFormat().setFontSize(14);
text.getCharacterFormat().setTextColor(new Color(255, 140, 0));

// Set the marker format of the endnote
endnote.getMarkerCharacterFormat().setFontName("Calibri");
endnote.getMarkerCharacterFormat().setFontSize(20);
endnote.getMarkerCharacterFormat().setTextColor(new Color(0, 0, 139));
```

---

# Spire.Doc Footnote Insertion
## Insert and format footnotes in a Word document
```java
// Find the text "Spire.Doc" in the document
TextSelection selection = document.findString("Spire.Doc", false, true);

// Get the TextRange
TextRange textRange = selection.getAsOneRange();

// Get the owner paragraph
Paragraph paragraph = textRange.getOwnerParagraph();

// Get the index of the paragraph
int index = paragraph.getChildObjects().indexOf(textRange);

// Append a footnote to the paragraph
Footnote footnote = paragraph.appendFootnote(FootnoteType.Footnote);

// Insert the footnote
paragraph.getChildObjects().insert(index + 1, footnote);

// Add text to the body of the footnote
textRange = footnote.getTextBody().addParagraph().appendText("Welcome to evaluate Spire.Doc");

// Set the format of the text in the footnote
textRange.getCharacterFormat().setFontName("Arial Black");
textRange.getCharacterFormat().setFontSize(10);
textRange.getCharacterFormat().setTextColor(new Color(255, 140, 0));

// Set the format of the footnote marker
footnote.getMarkerCharacterFormat().setFontName("Calibri");
footnote.getMarkerCharacterFormat().setFontSize(12);
footnote.getMarkerCharacterFormat().setBold(true);
footnote.getMarkerCharacterFormat().setTextColor(new Color(0, 0, 139));
```

---

# Remove Footnote from Document
## Remove footnotes from a Word document by traversing paragraphs
```java
// Get the first section from the document
Section section = document.getSections().get(0);

// Traverse through each paragraph in the section to find footnotes
for (int i = 0; i < section.getParagraphs().getCount(); i++) {
    Paragraph para = section.getParagraphs().get(i);
    int index = -1;

    // Iterate over child objects in the paragraph to find a Footnote
    for (int j = 0, cnt = para.getChildObjects().getCount(); j < cnt; j++) {
        ParagraphBase pBase = (ParagraphBase) para.getChildObjects().get(j);

        // Check if the current object is a Footnote
        if (pBase instanceof Footnote) {
            index = j;
            break;
        }
    }

    // If a Footnote is found, remove it from the paragraph's child objects
    if (index > -1)
        para.getChildObjects().removeAt(index);
}
```

---

# Spire.Doc Footnote Configuration
## Set footnote position, number format and restart rule
```java
// Get the first section
Section sec = doc.getSections().get(0);

// Set the number format
sec.getFootnoteOptions().setNumberFormat(FootnoteNumberFormat.Upper_Case_Letter);

// Set the restart rule
sec.getFootnoteOptions().setRestartRule(FootnoteRestartRule.Restart_Page);

// Set the position
sec.getFootnoteOptions().setPosition(FootnotePosition.Print_As_End_Of_Section);
```

---

# Spire.Doc Custom Paper Size
## Customize paper size for printing documents
```java
//Create a document
Document doc = new Document();

//Get the PrintDocument object
PrintDocument printDoc = doc.getPrintDocument();

//Custom the paper size
PaperSize size = new PaperSize();
size.setWidth(900);
size.setHeight(800);

//Apply the page size
printDoc.getDefaultPageSettings().setPaperSize(size);

//Print the document
printDoc.print();
```

---

# Spire.Doc Document Printing
## Print a Word document using Spire.Doc library
```java
// Create a document
Document document = new Document();

// Print the document
document.getPrintDocument().print();

//Dispose the document
document.dispose();
```

---

# Spire.Doc Document Printing
## Print a document with customized printer settings
```java
// Create a PrinterJob
PrinterJob loPrinterJob = PrinterJob.getPrinterJob();

// Get the default page format
PageFormat loPageFormat  = loPrinterJob.defaultPage();

// Get the paper
Paper loPaper = loPageFormat.getPaper();

// Delete the default margin
loPaper.setImageableArea(0,0,loPageFormat.getWidth(),loPageFormat.getHeight());
// Set the copy number
loPrinterJob.setCopies(1);

// Set the paper
loPageFormat.setPaper(loPaper);

// Enable printable
loPrinterJob.setPrintable(document,loPageFormat);

// Print the document
loPrinterJob.print();
```

---

# Spire.Doc Print Settings
## Set margins and duplex printing for document
```java
// Create
Document doc = new Document();
doc.loadFromFile("data/print.docx");

//Get the PrintDocument object
PrintDocument printDoc = doc.getPrintDocument();

//Set graphics origin starts at the page margins
printDoc.setOriginAtMargins(true);
//Set the margin to 0
printDoc.getDefaultPageSettings().setMargins(new Margins(0, 0, 0, 0));

//Double-sided, vertical printing
printDoc.getPrinterSettings().setDuplex(Duplex.Vertical);

//Double-sided, horizontal printing
//printDoc.getPrinterSettings().setDuplex(Duplex.Horizontal);

//Print the Word document
printDoc.print();
```

---

# Spire.Doc VBA Macro Detection and Removal
## Detect and remove VBA macros from Word documents
```java
// Create a Word document
Document document = new Document();

// If the document contains Macros, remove them from the document
if (document.isContainMacro()) {
    document.clearMacros();
}
```

---

# Document Macro Handling
## Load and save Word documents containing macros
```java
// Create a document
Document doc = new Document();

// Load Word document which contains macro
doc.loadFromFile("data/macros.docm", FileFormat.Docm);

// Save the file
String output = "output/macros.docm";
doc.saveToFile(output, FileFormat.Docm);

// Dispose the document
doc.dispose();
```

---

# Spire.Doc Picture Caption
## Add captions to pictures in a Word document
```java
//Create word document
Document document = new Document();

//Create a new section
Section section = document.addSection();

//Add a new paragraph
Paragraph par1 = section.addParagraph();

//Set the afters-pacing
par1.getFormat().setAfterSpacing(10);

//Load a picture
DocPicture pic1 = par1.appendPicture("data/spire.Doc.png");

//Set picture height
pic1.setHeight(120);

//Set picture width
pic1.setWidth(120);

//Create a CaptionNumberingFormat
CaptionNumberingFormat format = CaptionNumberingFormat.Number;

//Add caption to the picture
pic1.addCaption("Figure", format, CaptionPosition.Below_Item);

//Add the second paragraph
Paragraph par2 = section.addParagraph();

//Load a picture
DocPicture pic2 = par2.appendPicture("data/word.png");

//Set picture height
pic2.setHeight(120);

//Set picture width
pic2.setWidth(120);

//Add caption to the picture
pic2.addCaption("Figure", format, CaptionPosition.Below_Item);

//Update fields
document.isUpdateFields(true);
```

---

# Spire.Doc Table Caption
## Add caption to a table in Word document
```java
//Get the first table
Body body = document.getSections().get(0).getBody();

//Get the first document
Table table = body.getTables().get(0);

//Add caption to the table
table.addCaption("Table", CaptionNumberingFormat.Number, CaptionPosition.Below_Item);

//Update fields
document.isUpdateFields(true);
```

---

# Spire.Doc Picture Caption Cross-Reference
## Create cross-references for picture captions in Word documents
```java
//Add caption to the picture
CaptionNumberingFormat format = CaptionNumberingFormat.Number;
IParagraph captionParagraph = pic1.addCaption("Figure", format, CaptionPosition.Below_Item);

//Create a bookmark for the second picture caption
String bookmarkName = "Figure_2";
Paragraph paragraph = section.addParagraph();
paragraph.appendBookmarkStart(bookmarkName);
paragraph.appendBookmarkEnd(bookmarkName);

//Replace bookmark content with caption
BookmarksNavigator navigator = new BookmarksNavigator(document);
navigator.moveToBookmark(bookmarkName);
TextBodyPart part = navigator.getBookmarkContent();
part.getBodyItems().clear();
part.getBodyItems().add(captionParagraph);
navigator.replaceBookmarkContent(part);

//Create cross-reference field pointing to the bookmark
Field field = new Field(document);
field.setType(FieldType.Field_Ref);
field.setCode("REF Figure_2 \\p \\h");
firstPara.getChildObjects().add(field);
FieldMark fieldSeparator = new FieldMark(document, FieldMarkType.Field_Separator);
firstPara.getChildObjects().add(fieldSeparator);

//Set the display text of the field
TextRange tr = new TextRange(document);
tr.setText("Figure 2");
firstPara.getChildObjects().add(tr);

FieldMark fieldEnd = new FieldMark(document, FieldMarkType.Field_End);
firstPara.getChildObjects().add(fieldEnd);

//Update fields
document.isUpdateFields(true);
```

---

# Table Caption Cross Reference
## Create table caption with cross-reference field in Spire.Doc
```java
//Create word document
Document document = new Document();

//Add a new section
Section section = document.addSection();

//Create a table
Table table = section.addTable(true);

//Set the number of rows and columns
table.resetCells(2, 3);

//Add caption to the table
IParagraph captionParagraph = table.addCaption("Table", CaptionNumberingFormat.Number, CaptionPosition.Below_Item);

//Create a bookmark
String bookmarkName = "Table_1";

//Add a paragraph
Paragraph paragraph = section.addParagraph();

//Append the bookmark
paragraph.appendBookmarkStart(bookmarkName);
paragraph.appendBookmarkEnd(bookmarkName);

//Create a BookmarksNavigator
BookmarksNavigator navigator = new BookmarksNavigator(document);

//Move the navigator to the bookmark
navigator.moveToBookmark(bookmarkName);
TextBodyPart part = navigator.getBookmarkContent();
part.getBodyItems().clear();
part.getBodyItems().add(captionParagraph);

//Replace bookmark content
navigator.replaceBookmarkContent(part);

//Create cross-reference field to point to bookmark "Table_1"
Field field = new Field(document);
field.setType(FieldType.Field_Ref);
field.setCode("REF Table_1 \\p \\h");

//Insert line breaks
for (int i = 0; i < 3; i++) {
    paragraph.appendBreak(BreakType.Line_Break);
}

//Insert field to paragraph
paragraph = section.addParagraph();
TextRange range = paragraph.appendText("This is a table caption cross-reference, ");
range.getCharacterFormat().setFontSize(14);
paragraph.getChildObjects().add(field);

//Insert FieldSeparator object
FieldMark fieldSeparator = new FieldMark(document, FieldMarkType.Field_Separator);
paragraph.getChildObjects().add(fieldSeparator);

//Set display text of the field
TextRange tr = new TextRange(document);
tr.setText("Table 1");
tr.getCharacterFormat().setFontSize(14);
tr.getCharacterFormat().setTextColor(Color.cyan);
paragraph.getChildObjects().add(tr);

//Insert FieldEnd object to mark the end of the field
FieldMark fieldEnd = new FieldMark(document, FieldMarkType.Field_End);
paragraph.getChildObjects().add(fieldEnd);

//Update fields
document.isUpdateFields(true);
```

---

# Spire.Doc Bar Chart Creation
## Append a bar chart to a Word document
```java
// Create a new instance of Document
Document document = new Document();

// Add a section to the document
Section section = document.addSection();

// Add a paragraph to the section and append text to it
section.addParagraph().appendText("Bar chart.");

// Add a new paragraph to the section
Paragraph newPara = section.addParagraph();

// Append a bar chart shape to the paragraph with specified width and height
ShapeObject chartShape = newPara.appendChart(ChartType.Bar, 400, 300);
Chart chart = chartShape.getChart();

// Get the title of the chart
ChartTitle title = chart.getTitle();

// Set the text of the chart title
title.setText("My Chart");

// Show the chart title
title.setShow(true);

// Overlay the chart title on top of the chart
title.setOverlay(true);
```

---

# Spire.Doc Bubble Chart
## Append a bubble chart to a Word document
```java
// Create a new instance of Document
Document document = new Document();

// Add a section to the document
Section section = document.addSection();

// Add a paragraph to the section and append text to it
section.addParagraph().appendText("Bubble chart.");

// Add a new paragraph to the section
Paragraph newPara = section.addParagraph();

// Append a bubble chart shape to the paragraph with specified width and height
ShapeObject shape = newPara.appendChart(ChartType.Bubble, 500, 300);

// Get the chart object from the shape
Chart chart = shape.getChart();

// Clear any existing series in the chart
chart.getSeries().clear();

// Add a new series to the chart with data points for X, Y, and bubble size values
ChartSeries series = chart.getSeries().add("E-iceblue Test Series",
        new double[] { 2.9, 3.5, 1.1, 4.0, 4.0 },
        new double[] { 1.9, 8.5, 2.1, 6.0, 1.5 },
        new double[] { 9.0, 4.5, 2.5, 8.0, 5.0 });
```

---

# Spire.Doc Column Chart
## Append a column chart to a Word document
```java
// Create a new instance of Document
Document document = new Document();

// Add a section to the document
Section section = document.addSection();

// Add a new paragraph to the section
Paragraph newPara = section.addParagraph();

// Append a column chart shape to the paragraph with specified width and height
ShapeObject shape = newPara.appendChart(ChartType.Column, 500, 300);

// Get the chart object from the shape
Chart chart = shape.getChart();

// Clear any existing series in the chart
chart.getSeries().clear();

// Add a new series to the chart with data points for X values (categories) and Y values
chart.getSeries().add("E-iceblue Test Series",
        new String[] { "Word", "PDF", "Excel", "GoogleDocs", "Office" },
        new double[] { 1900000, 850000, 2100000, 600000, 1500000 });

// Set the number format for the Y-axis labels
chart.getAxisY().getNumberFormat().setFormatCode("#,##0");
```

---

# Spire.Doc Line Chart Creation
## Create and append a line chart to a Word document with custom data series
```java
// Add a new paragraph to the section
Paragraph newPara = section.addParagraph();

// Append a line chart shape to the paragraph with specified width and height
ShapeObject shape = newPara.appendChart(ChartType.Line, 500, 300);

// Get the chart object from the shape
Chart chart = shape.getChart();

// Get the title of the chart
ChartTitle title = chart.getTitle();

// Set the text of the chart title
title.setText("My Chart");

// Clear any existing series in the chart
ChartSeriesCollection seriesColl = chart.getSeries();
seriesColl.clear();

// Define categories (X-axis values)
String[] categories = { "C1", "C2", "C3", "C4", "C5", "C6" };

// Add two series to the chart with specified categories and Y-axis values
seriesColl.add("AW Series 1", categories, new double[] { 1, 2, 2.5, 4, 5, 6 });
seriesColl.add("AW Series 2", categories, new double[] { 2, 3, 3.5, 6, 6.5, 7 });
```

---

# spire.doc pie chart
## append pie chart to document with data series
```java
// Append a pie chart shape to the paragraph with specified width and height
ShapeObject shape = newPara.appendChart(ChartType.Pie, 500, 300);
Chart chart = shape.getChart();

// Add a series to the chart with categories (labels) and corresponding data values
ChartSeries series = chart.getSeries().add("E-iceblue Test Series",
        new String[] { "Word", "PDF", "Excel" },
        new double[] { 2.7, 3.2, 0.8 });
```

---

# spire.doc scatter chart
## append scatter chart to Word document
```java
// Create a new instance of Document
Document document = new Document();

// Add a section to the document
Section section = document.addSection();

// Add a paragraph to the section and append text to it
section.addParagraph().appendText("Scatter chart.");

// Add a new paragraph to the section
Paragraph newPara = section.addParagraph();

// Append a scatter chart shape to the paragraph with specified width and height
ShapeObject shape = newPara.appendChart(ChartType.Scatter, 450, 300);
Chart chart = shape.getChart();
// Clear any existing series in the chart
chart.getSeries().clear();

// Add a new series to the chart with data points for X and Y values
chart.getSeries().add("Series 1",
        new double[] { 1.0, 2.0, 3.0, 4.0, 5.0 },
        new double[] { 1.0, 20.0, 40.0, 80.0, 160.0 });
```

---

# Spire.Doc 3D Surface Chart
## Create and configure a 3D surface chart in a Word document
```java
// Append a Surface3D chart shape to the paragraph with specified width and height
ShapeObject shape = newPara.appendChart(ChartType.Surface_3_D, 500, 300);

// Get the chart object from the shape
Chart chart = shape.getChart();
// Clear any existing series in the chart
chart.getSeries().clear();

// Set the title of the chart
chart.getTitle().setText("My chart");

// Add multiple series to the chart with categories (X-axis values) and corresponding data values
chart.getSeries().add("E-iceblue Test Series 1",
        new String[] { "Word", "PDF", "Excel", "GoogleDocs", "Office" },
        new double[] { 1900000, 850000, 2100000, 600000, 1500000 });

chart.getSeries().add("E-iceblue Test Series 2",
        new String[] { "Word", "PDF", "Excel", "GoogleDocs", "Office" },
        new double[] { 900000, 50000, 1100000, 400000, 2500000 });

chart.getSeries().add("E-iceblue Test Series 3",
        new String[] { "Word", "PDF", "Excel", "GoogleDocs", "Office" },
        new double[] { 500000, 820000, 1500000, 400000, 100000 });
```

---

# Spire.Doc Fixed Layout Analysis
## Extract layout information from a document including lines, paragraphs, and page details
```java
// Create a FixedLayoutDocument object using the loaded document
FixedLayoutDocument layoutDoc = new FixedLayoutDocument(doc);

// Get the first line on the first page
FixedLayoutLine line = layoutDoc.getPages().get(0).getColumns().get(0).getLines().get(0);

// Retrieve the original paragraph associated with the line
Paragraph para = line.getParagraph();

// Retrieve all the text on the first page, including headers and footers
String pageText = layoutDoc.getPages().get(0).getText();

// Iterate through each page in the document to get the number of lines on each page
for (Object obj : layoutDoc.getPages()) {
    FixedLayoutPage page = (FixedLayoutPage) obj;
    LayoutCollection<LayoutElement> lines = page.getChildEntities(LayoutElementType.Line, true);
}

// Perform a reverse lookup of layout entities for the first paragraph
for (Object object : layoutDoc.getLayoutEntitiesOfNode(((Section) doc.getFirstChild()).getBody().getParagraphs().get(0))) {
    FixedLayoutLine paragraphLine = (FixedLayoutLine) object;
    paragraphLine.getText().trim();
    paragraphLine.getRectangle().toString();
}
```

---

# Spire.Doc Paragraph Formatting
## Remove space between paragraphs of same style
```java
// Get the Body object from the first section of the document
Body body = document.getSections().get(0).getBody();

// Loop through each paragraph in the body of the document
Paragraph paragraph;
for (int i = 0; i < body.getParagraphs().getCount(); i++) {
    // Retrieve the current paragraph
    paragraph = body.getParagraphs().get(i);

    // Set no space between paragraphs of the same style for the current paragraph
    paragraph.getFormat().setNoSpaceBetweenParagraphsOfSameStyle(true);
}
```

---

# Spire.Doc Document Comparison
## Compare documents ignoring table differences
```java
// Create a new CompareOptions object to specify comparison settings
CompareOptions compareoptions = new CompareOptions();

// Set the option to ignore differences in tables during comparison
compareoptions.setIgnoreTable(true);

// Compare the two documents using the specified options, with "E-iceblue" as the author name for tracked changes
document1.compare(document2, "E-iceblue", compareoptions);
```

---

# Spire.Doc Document Revision Tracking
## Track document revisions in Word documents
```java
// Start the track revisions
document.startTrackRevisions("User01", new Date());

// Get the first paragraph and add content
document.getSections().get(0).getParagraphs().get(0).appendText("User01 add new Text!");

// Delete a paragraph
document.getSections().get(0).getParagraphs().removeAt(2);

// Stop the track revisions
document.stopTrackRevisions();
```

---

# Spire.Doc document conversion
## Convert Word document to MHTML format
```java
// Create word document
Document document = new Document();

// Load the file from disk
document.loadFromFile("data\\ToMhtml.docx");

// Save to MHTML file
document.saveToFile("ToMhtml-out.mhtml", FileFormat.Mhtml);
```

---

# Spire.Doc Java Text Formatting
## Set underline color and style for text in document
```java
// Add a new paragraph to the section
Paragraph paragraph = section.addParagraph();

// Append text to the paragraph and get the TextRange object for formatting
TextRange textRange = paragraph.appendText("Welcome to evaluate Spire.Doc for Java product.");

// Set the underline style of the text to single underline
textRange.getCharacterFormat().setUnderlineStyle(UnderlineStyle.Single);

// Set the underline color of the text to red
textRange.getCharacterFormat().setUnderlineColor(Color.red);
```

---

# Spire.Doc Bookmark Modification
## Modify bookmark name in Word document
```java
// Create a new instance of the Document class
Document document = new Document();

// Retrieve the Bookmark object
Bookmark bookmark = document.getBookmarks().get("Test");

// Change the name of the retrieved bookmark to "bookmark1"
bookmark.setName("bookmark1");
```

---

# Spire.Doc Chart Axis Configuration
## Configure chart axes properties including labels, gridlines, titles, and formatting
```java
static void appendChartAxis(Chart chart)
{
    for (int i = 0; i < chart.getAxes().getCount(); i++)
    {
        if (i == 0)
        {
            chart.getAxes().get(i).setCategoryType(AxisCategoryType.Category);
            chart.getAxes().get(i).getBounds().setMaximum(new AxisBound(5));
            chart.getAxes().get(i).getBounds().setMinimum(new AxisBound(0));
            chart.getAxes().get(i).getUnits().setMajor(1);
            chart.getAxes().get(i).getUnits().setMajorTimeUnit(AxisTimeUnit.Auto);
            chart.getAxes().get(i).getUnits().setMinor(1);
            chart.getAxes().get(i).getUnits().setMinorTimeUnit(AxisTimeUnit.Days);
            chart.getAxes().get(i).hasMajorGridlines(false);
            chart.getAxes().get(i).hasMinorGridlines(true);
            chart.getAxes().get(i).getLabels().isAutoSpacing(false);
            chart.getAxes().get(i).getLabels().setSpacing(1);
            chart.getAxes().get(i).getLabels().setOffset(1);
            chart.getAxes().get(i).getLabels().setPosition(AxisTickLabelPosition.Low);
            chart.getAxes().get(i).setReverseOrder(true);
            chart.getAxes().get(i).getTitle().setText("x-axis");
            chart.getAxes().get(i).getTitle().setShow(true);
            chart.getAxes().get(i).getTitle().setOverlay(true);
        }
        else if (i == 1)
        {
            chart.getAxes().get(i).setCategoryType(AxisCategoryType.Category);
            chart.getAxes().get(i).getUnits().isMajorAuto(true);
            chart.getAxes().get(i).getUnits().isMinorAuto(true);
            chart.getAxes().get(i).getBounds().setLogBase(10);
            chart.getAxes().get(i).hasMajorGridlines(true);
            chart.getAxes().get(i).hasMinorGridlines(false);
            chart.getAxes().get(i).setReverseOrder(false);
            chart.getAxes().get(i).getLabels().isAutoSpacing(true);
            chart.getAxes().get(i).getTitle().setText("y-axis");
            chart.getAxes().get(i).getTitle().setShow(true);
            chart.getAxes().get(i).getTitle().setOverlay(true);
        }
        else
        {
            chart.getAxes().get(i).getTitle().setText("z-axis");
            chart.getAxes().get(i).getTitle().setShow(true);
            chart.getAxes().get(i).getTitle().setOverlay(false);
        }
        chart.getAxes().get(i).getLabels().setAlignment(LabelAlignment.Left);
        chart.getAxes().get(i).getUnits().setBaseTimeUnit(AxisTimeUnit.Auto);
        chart.getAxes().get(i).setAxisBetweenCategories(true);
        chart.getAxes().get(i).getDisplayUnits().setCustomUnit(1);
        chart.getAxes().get(i).getDisplayUnits().setUnit(AxisBuiltInUnit.Custom);
        chart.getAxes().get(i).getDisplayUnits().setShowLabel(true);
        chart.getAxes().get(i).getTickMarks().setSpacing(1);
        chart.getAxes().get(i).getTickMarks().setMajor(AxisTickMark.None);
        chart.getAxes().get(i).getTickMarks().setMinor(AxisTickMark.Inside);
        chart.getAxes().get(i).getTitle().getCharacterFormat().setFontSize(8);
        chart.getAxes().get(i).getTitle().getCharacterFormat().setTextColor(Color.red);
        chart.getAxes().get(i).getTitle().getCharacterFormat().setBold(true);
    }
}
```

---

# Spire.Doc Chart Data Labels
## Find charts in document and configure data labels
```java
// Loop through all sections in the document
for (int i = 0; i < document.getSections().getCount(); i++) {
    // Loop through all paragraphs in the current section
    for (int j = 0; j < document.getSections().get(i).getParagraphs().getCount(); j++) {
        // Get the current paragraph
        Paragraph paragraph = document.getSections().get(i).getParagraphs().get(j);

        // Loop through all child objects in the paragraph
        for (int k = 0; k < paragraph.getChildObjects().getCount(); k++) {
            DocumentObject obj = paragraph.getChildObjects().get(k);
            // Check if the object is a shape (e.g., chart, etc.)
            if (obj instanceof ShapeObject) {
                // Cast the object to a ShapeObject
                ShapeObject shape = (ShapeObject) obj;

                // Get the chart from the shape
                Chart chart = shape.getChart();
                ChartSeriesCollection series = chart.getSeries();
                ChartDataLabelCollection dataLabels = series.get(0).getDataLabels();
                series.get(0).hasDataLabels(true);
                
                // Configure data labels
                dataLabels.setShowValue(true);
                dataLabels.setShowCategoryName(true);
                dataLabels.setShowSeriesName(true);
                dataLabels.setShowLeaderLines(true);
                dataLabels.setSeparator(";");
                dataLabels.getNumberFormat().setFormatCode("#,##0");

                // Configure text formatting
                dataLabels.getCharacterFormat().setFontSize(8);
                dataLabels.getCharacterFormat().setBold(true);
                dataLabels.getCharacterFormat().setTextColor(Color.blue);
                dataLabels.getCharacterFormat().getBorder().setColor(Color.blue);
                dataLabels.getCharacterFormat().setBidi(true);
                dataLabels.getCharacterFormat().setItalic(true);
                dataLabels.getCharacterFormat().setUnderlineColor(Color.red);
                dataLabels.getCharacterFormat().setUnderlineStyle(UnderlineStyle.Double);
                dataLabels.getCharacterFormat().setFontName("Arial");
                dataLabels.getCharacterFormat().setAllCaps(true);
                dataLabels.getCharacterFormat().isShadow(true);
            }
        }
    }
}
```

---

# Spire.Doc chart data table
## append chart data table to Word document
```java
static void appendChartDataTable(Chart chart)
{
    // Enable the display of the data table in the chart
    chart.getDataTable().setShow(true);

    // Show legend keys (symbols) in the data table
    chart.getDataTable().setShowLegendKeys(true);

    // Display horizontal borders between rows in the data table
    chart.getDataTable().setShowHorizontalBorder(true);

    // Display vertical borders between columns in the data table
    chart.getDataTable().setShowVerticalBorder(true);

    // Show an outline border around the entire data table
    chart.getDataTable().setShowOutlineBorder(true);
}
```

---

