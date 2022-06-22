import com.spire.doc.*;
import com.spire.doc.documents.*;
import com.spire.doc.fields.*;

import java.awt.*;

public class addRichTextContentControl {
    public static void main(String[] args) {
        //Create a document
        Document document = new Document();

        //Add a new section.
        Section section = document.addSection();

        //Add a paragraph
        Paragraph paragraph = section.addParagraph();

        //Append textRange for the paragraph
        TextRange txtRange = paragraph.appendText("The following example shows how to add RichText content control in a Word document. \n");

        //Append textRange
        txtRange = paragraph.appendText("Add RichText Content Control:  ");

        //Set the font format
        txtRange.getCharacterFormat().setItalic(true);

        //Create StructureDocumentTagInline for document
        StructureDocumentTagInline sdt = new StructureDocumentTagInline(document);

        //Add sdt in paragraph
        paragraph.getChildObjects().add(sdt);

        //Specify the type
        sdt.getSDTProperties().setSDTType(SdtType.Rich_Text);

        //Set displaying text
        SdtText text = new SdtText(true);
        text.isMultiline(true);
        sdt.getSDTProperties().setControlProperties(text);

        //Crate a TextRange
        TextRange rt = new TextRange(document);
        rt.setText("Welcome to use ");
        rt.getCharacterFormat().setTextColor(Color.GREEN);
        sdt.getSDTContent().getChildObjects().add(rt);

        rt = new TextRange(document);
        rt.setText("Spire.Doc");
        rt.getCharacterFormat().setTextColor(Color.ORANGE);
        sdt.getSDTContent().getChildObjects().add(rt);

        //Save the document
        String output = "output/addRichTextContentControl.docx";
        document.saveToFile(output, FileFormat.Docx);
    }
}
