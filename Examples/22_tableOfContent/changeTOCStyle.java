import com.spire.doc.*;
import com.spire.doc.documents.*;

import java.awt.*;

public class changeTOCStyle {
    public static void main(String[] args) {
        //Load document from disk
        Document doc = new Document();
        doc.loadFromFile("data/template_Toc.docx");

        //Defind a Toc style
        ParagraphStyle tocStyle = (ParagraphStyle) Style.createBuiltinStyle(BuiltinStyle.Toc_1, doc);
        tocStyle.getCharacterFormat().setFontName("Aleo");
        tocStyle.getCharacterFormat().setFontSize(15f);
        tocStyle.getCharacterFormat().setTextColor(Color.LIGHT_GRAY);
        doc.getStyles().add(tocStyle);


        //Loop through sections
        for (int s = 0; s < doc.getSections().getCount(); s++) {
            Section section = doc.getSections().get(s);
            //Loop through content of section
            for (int i = 0; i < section.getBody().getChildObjects().getCount(); i++) {
                DocumentObject obj = section.getBody().getChildObjects().get((i));
                //Find the structure document tag
                if (obj instanceof StructureDocumentTag) {
                    StructureDocumentTag tag = (StructureDocumentTag) ((obj instanceof StructureDocumentTag) ? obj : null);
                    //Find the paragraph where the TOC1 locates
                    for (int j = 0; j < tag.getChildObjects().getCount(); j++) {
                        DocumentObject cObj = tag.getChildObjects().get(j);
                        if (cObj instanceof Paragraph) {
                            Paragraph para = (Paragraph) ((cObj instanceof Paragraph) ? cObj : null);
                            if (para.getStyleName().equals("TOC1")) {
                                //Apply the new style for TOC1 paragraph
                                para.applyStyle(tocStyle.getName());
                            }
                        }
                    }
                }
            }
        }

        //Save the Word file
        String output = "output/changeTOCStyle.docx";
        doc.saveToFile(output, FileFormat.Docx_2013);
    }
}
