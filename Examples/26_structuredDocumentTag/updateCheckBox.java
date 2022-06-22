import com.spire.doc.*;
import com.spire.doc.documents.*;

import java.util.*;

public class updateCheckBox {
    public static void main(String[] args) {
        //Create a document
        Document document = new Document();

        //Load the document from disk.
        document.loadFromFile("data/checkBoxContentControl.docx");

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
                SdtCheckBox scb = (SdtCheckBox)tagInlines.get(i).getSDTProperties().getControlProperties() ;
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
        //Save the document
        String output = "output/updateCheckBox.docx";
        document.saveToFile(output, FileFormat.Docx);
    }

    public static StructureTagInLines GetAllTags(Document document) {
        StructureTagInLines structureTags = new StructureTagInLines();
        for (int i = 0; i < document.getSections().getCount(); i++) {
            Section section = document.getSections().get(i);
            for (int j = 0; j < section.getBody().getChildObjects().getCount(); j++){
                DocumentObject obj = section.getBody().getChildObjects().get(j);
                if(obj.getDocumentObjectType() == DocumentObjectType.Paragraph)
                    for (int k = 0; k < ((Paragraph)obj).getChildObjects().getCount() ; k++) {
                        DocumentObject pobj = ((Paragraph)obj).getChildObjects().get(k);
                        if (pobj.getDocumentObjectType() == DocumentObjectType.Structure_Document_Tag_Inline) {
                            structureTags.getM_tagInlines().add((StructureDocumentTagInline)pobj);
                        }
                    }

            }
        }
        return structureTags;
    }
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
