import com.spire.doc.*;
import com.spire.doc.documents.*;
import com.spire.doc.fields.*;

import java.util.*;

public class addContentControls {
    public static void main(String[] args) {
        //Create a new Word document
        Document document = new Document();
        Section section = document.addSection();
        Paragraph paragraph = section.addParagraph();
        TextRange txtRange = paragraph.appendText("The following example shows how to add content controls in a Word document.");
        section.addParagraph();

        //Add combo box content control
        paragraph = section.addParagraph();
        txtRange = paragraph.appendText("Add Combo Box Content Control:  ");
        txtRange.getCharacterFormat().setItalic(true);
        StructureDocumentTagInline sd = new StructureDocumentTagInline(document);
        paragraph.getChildObjects().add(sd);
        sd.getSDTProperties().setSDTType(SdtType.Combo_Box);
        SdtComboBox cb = new SdtComboBox();
        cb.getListItems().add(new SdtListItem("Spire.Doc"));
        cb.getListItems().add(new SdtListItem("Spire.XLS"));
        cb.getListItems().add(new SdtListItem("Spire.PDF"));
        sd.getSDTProperties().setControlProperties(cb);
        TextRange rt = new TextRange(document);
        rt.setText(cb.getListItems().get(0).getDisplayText());
        sd.getSDTContent().getChildObjects().add(rt);
        section.addParagraph();

        //Add text content control
        paragraph = section.addParagraph();
        txtRange = paragraph.appendText("Add Text Content Control:  ");
        txtRange.getCharacterFormat().setItalic(true);
        sd = new StructureDocumentTagInline(document);
        paragraph.getChildObjects().add(sd);
        sd.getSDTProperties().setSDTType(SdtType.Text);
        SdtText text = new SdtText(true);
        text.isMultiline(true);
        sd.getSDTProperties().setControlProperties(text);
        rt = new TextRange(document);
        rt.setText("Text");
        sd.getSDTContent().getChildObjects().add(rt);
        section.addParagraph();

        //Add picture content control
        paragraph = section.addParagraph();
        txtRange = paragraph.appendText("Add Picture Content Control:  ");
        txtRange.getCharacterFormat().setItalic(true);
        sd = new StructureDocumentTagInline(document);
        paragraph.getChildObjects().add(sd);
        sd.getSDTProperties().setSDTType(SdtType.Picture);
        DocPicture pic = new DocPicture(document);
        pic.setWidth(10f);
        pic.setHeight(10f);
        pic.loadImage("data/logo.png");
        sd.getSDTContent().getChildObjects().add(pic);
        section.addParagraph();

        //Add date picker content control
        paragraph = section.addParagraph();
        txtRange = paragraph.appendText("Add Date Picker Content Control:  ");
        txtRange.getCharacterFormat().setItalic(true);
        sd = new StructureDocumentTagInline(document);
        paragraph.getChildObjects().add(sd);
        sd.getSDTProperties().setSDTType(SdtType.Date_Picker);
        SdtDate date = new SdtDate();
        date.setCalendarType(CalendarType.Default);
        date.setDateFormat("yyyy.MM.dd");
        date.setFullDate(new Date());
        sd.getSDTProperties().setControlProperties(date);
        rt = new TextRange(document);
        rt.setText("2019.12.31");
        sd.getSDTContent().getChildObjects().add(rt);
        section.addParagraph();

        //Add drop-down list content control
        paragraph = section.addParagraph();
        txtRange = paragraph.appendText("Add Drop-Down List Content Control:  ");
        txtRange.getCharacterFormat().setItalic(true);
        sd = new StructureDocumentTagInline(document);
        paragraph.getChildObjects().add(sd);
        sd.getSDTProperties().setSDTType(SdtType.Drop_Down_List);
        SdtDropDownList sddl = new SdtDropDownList();
        sddl.getListItems().add(new SdtListItem("Harry"));
        sddl.getListItems().add(new SdtListItem("Jerry"));
        sd.getSDTProperties().setControlProperties(sddl);
        rt = new TextRange(document);
        rt.setText(sddl.getListItems().get(0).getDisplayText());
        sd.getSDTContent().getChildObjects().add(rt);

        //Save to file
        String output = "output/addContentControls.docx";
        document.saveToFile(output, FileFormat.Docx);
    }
}
