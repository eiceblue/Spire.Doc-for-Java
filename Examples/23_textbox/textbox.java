import com.spire.doc.*;
import com.spire.doc.documents.*;
import com.spire.doc.fields.*;

import java.awt.*;

public class textbox {
    public static void main(String[] args) {
        //Create a Word document and and a section
        Document document = new Document();
        Section section = document.addSection();

        insertTextbox(section);

        //Save docx file
        String output = "output/textbox.docx";
        document.saveToFile(output, FileFormat.Docx_2013);
    }

    public static void insertTextbox(Section section) {
        section.addParagraph();
        Paragraph paragraph = section.addParagraph();

        //Insert and getFormat() the first textbox
        TextBox textBox1 = paragraph.appendTextBox(240, 35);
        textBox1.getFormat().setHorizontalAlignment(ShapeHorizontalAlignment.Left);
        textBox1.getFormat().setLineColor(Color.GRAY);
        textBox1.getFormat().setLineStyle(TextBoxLineStyle.Simple);
        textBox1.getFormat().setFillColor(Color.GREEN);
        Paragraph para = textBox1.getBody().addParagraph();
        TextRange txtrg = para.appendText("Textbox 1 in the document");
        txtrg.getCharacterFormat().setFontName("Lucida Sans Unicode");
        txtrg.getCharacterFormat().setFontSize(14);
        txtrg.getCharacterFormat().setTextColor(Color.white);
        para.getFormat().setHorizontalAlignment(HorizontalAlignment.Center);

        //Insert and getFormat() the second textbox
        section.addParagraph();
        section.addParagraph();
        section.addParagraph();
        paragraph = section.addParagraph();
        TextBox textBox2 = paragraph.appendTextBox(240, 35);
        textBox2.getFormat().setHorizontalAlignment(ShapeHorizontalAlignment.Left);
        textBox2.getFormat().setLineColor(Color.red);
        textBox2.getFormat().setLineStyle(TextBoxLineStyle.Thick_Thin);
        textBox2.getFormat().setFillColor(Color.blue);
        textBox2.getFormat().setLineDashing(LineDashing.Dot);
        para = textBox2.getBody().addParagraph();
        txtrg = para.appendText("Textbox 2 in the document");
        txtrg.getCharacterFormat().setFontName("Lucida Sans Unicode");
        txtrg.getCharacterFormat().setFontSize(14);
        txtrg.getCharacterFormat().setTextColor(Color.pink);
        para.getFormat().setHorizontalAlignment(HorizontalAlignment.Center);

        //Insert and getFormat() the third textbox
        section.addParagraph();
        section.addParagraph();
        section.addParagraph();
        paragraph = section.addParagraph();
        TextBox textBox3 = paragraph.appendTextBox(240, 35);
        textBox3.getFormat().setHorizontalAlignment(ShapeHorizontalAlignment.Left);
        textBox3.getFormat().setLineColor(Color.MAGENTA);
        textBox3.getFormat().setLineStyle(TextBoxLineStyle.Triple);
        textBox3.getFormat().setFillColor(Color.pink);
        textBox3.getFormat().setLineDashing(LineDashing.Dash_Dot_Dot);
        para = textBox3.getBody().addParagraph();
        txtrg = para.appendText("Textbox 3 in the document");
        txtrg.getCharacterFormat().setFontName("Lucida Sans Unicode");
        txtrg.getCharacterFormat().setFontSize(14);
        txtrg.getCharacterFormat().setTextColor(Color.red);
        para.getFormat().setHorizontalAlignment(HorizontalAlignment.Center);
    }
}
