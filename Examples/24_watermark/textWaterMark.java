import com.spire.doc.*;
import com.spire.doc.documents.*;

import java.awt.*;

public class textWaterMark {
    public static void main(String[] args) {
        //Open a Word document as template
        Document document = new Document("data/sampleB_2.docx");

        //Insert text watermark
        insertTextWatermark(document.getSections().get(0));

        //Save as docx file
        String output = "output/textWaterMark.docx";
        document.saveToFile(output, FileFormat.Docx);
    }

    public static void insertTextWatermark(Section section) {
        TextWatermark txtWatermark = new TextWatermark();
        txtWatermark.setText("E-iceblue");
        txtWatermark.setFontSize(25);
        txtWatermark.setColor(Color.blue);
        txtWatermark.setLayout(WatermarkLayout.Diagonal);
        section.getDocument().setWatermark(txtWatermark);

    }
}
