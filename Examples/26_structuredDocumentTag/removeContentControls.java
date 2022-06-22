import com.spire.doc.*;
import com.spire.doc.documents.*;

public class removeContentControls {
    public static void main(String[] args) {
        //Load document from disk
        Document doc = new Document();
        doc.loadFromFile("data/removeContentControls.docx");

        //Loop through sections
        for (int s = 0; s < doc.getSections().getCount(); s++) {
            Section section = doc.getSections().get(s);
            for (int i = 0; i < section.getBody().getChildObjects().getCount(); i++) {
                //Loop through contents in paragraph
                if (section.getBody().getChildObjects().get(i) instanceof Paragraph) {
                    Paragraph para = (Paragraph) section.getBody().getChildObjects().get(i);
                    for (int j = 0; j < para.getChildObjects().getCount(); j++) {
                        //Find the StructureDocumentTagInline
                        if (para.getChildObjects().get(j) instanceof StructureDocumentTagInline) {
                            com.spire.doc.documents.StructureDocumentTagInline sdt = (StructureDocumentTagInline) para.getChildObjects().get(j);
                            //Remove the content control from paragraph
                            para.getChildObjects().remove(sdt);
                            j--;
                        }
                    }
                }
                if (section.getBody().getChildObjects().get(i) instanceof StructureDocumentTag) {
                    com.spire.doc.documents.StructureDocumentTag sdt = (StructureDocumentTag) section.getBody().getChildObjects().get(i);
                    section.getBody().getChildObjects().remove(sdt);
                    i--;
                }
            }
        }

        //Save the Word document
        String output = "output/removeContentControls.docx";
        doc.saveToFile(output, FileFormat.Docx_2013);
    }
}
