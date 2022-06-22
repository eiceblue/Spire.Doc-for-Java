import com.spire.doc.*;
import com.spire.doc.documents.*;

public class changeTOCTabStyle {
    public static void main(String[] args) {
        //Load document from disk
        Document doc = new Document();
        doc.loadFromFile("data/template_Toc.docx");

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
                        {
                            Paragraph para = (Paragraph) cObj;
                            if (para.getStyleName().equals("TOC2")) {
                                //Set the tab style of paragraph
                                for (int t = 0; t < para.getFormat().getTabs().getCount(); t++) {
                                    Tab tab = para.getFormat().getTabs().get(t);
                                    tab.setPosition(tab.getPosition() + 20);
                                    tab.setTabLeader(TabLeader.No_Leader);
                                }
                            }
                        }
                    }
                }
            }
        }

        //Save the Word file
        String output = "output/changeTOCTabStyle.docx";
        doc.saveToFile(output, FileFormat.Docx_2013);
    }
}
