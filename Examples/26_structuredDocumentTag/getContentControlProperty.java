import com.spire.doc.*;
import com.spire.doc.documents.*;
import com.spire.doc.fields.TextRange;

import java.io.*;
import java.math.*;
import java.util.*;

public class getContentControlProperty {
    public static void main(String[] args) throws IOException {
        //Create a new document and load from file
        Document doc = new Document();
        doc.loadFromFile("data/contentControl.docx");
        String alias;
        BigDecimal id;
        String tag;
        String sdtType;
        SdtType sdt = SdtType.Rich_Text;
        String content = "";
        //Get all structureTags in the Word document
        StructureTags structureTags = GetAllTags(doc);

        String property = "Alias of contentControl" + "\t" + "ID          " + "\t" + "Tag             " + "\t" + "STDType        " +"\t"+"Content        " +  "\r\n";
        Paragraph paragraph;
        TextRange textRange;
        List<StructureDocumentTagInline> tagInlines = structureTags.getM_tagInlines();
        for (int i = 0; i < tagInlines.size(); i++)
        {
            alias = tagInlines.get(i).getSDTProperties().getAlias();
            id = tagInlines.get(i).getSDTProperties().getId();
            tag = tagInlines.get(i).getSDTProperties().getTag();
            sdt = tagInlines.get(i).getSDTProperties().getSDTType();
            sdtType = sdt.toString();
            if (sdt == SdtType.Rich_Text || sdt == SdtType.Text)
            {
                if (tagInlines.get(i).getChildObjects().getCount() > 0)
                {
                    for (int k = 0; k <  tagInlines.get(i).getChildObjects().getCount(); k++) {
                        if (tagInlines.get(i).getChildObjects().get(k).getDocumentObjectType() == DocumentObjectType.Text_Range)
                        {
                            textRange = (TextRange)tagInlines.get(i).getChildObjects().get(k);
                            content += textRange.getText();
                        }
                    }
                }
            }
            property += alias + ",\t" + id + ",\t" + tag + ",\t" + sdtType + ",\t" + content + "\r\n";
            content="";
        }


        List<StructureDocumentTag> tags = structureTags.getM_tags();
        for (int i = 0; i < tags.size(); i++) {
             alias = tags.get(i).getSDTProperties().getAlias();
             id = tags.get(i).getSDTProperties().getId();
             tag = tags.get(i).getSDTProperties().getTag();
            sdt = tags.get(i).getSDTProperties().getSDTType();
            sdtType = sdt.toString();
            if (sdt == SdtType.Rich_Text || sdt == SdtType.Text)
            {
                if (tags.get(i).getChildObjects().getCount() > 0)
                {
                    for (int k = 0; k <  tags.get(i).getChildObjects().getCount(); k++) {
                        if (tags.get(i).getChildObjects().get(k).getDocumentObjectType() == DocumentObjectType.Paragraph)
                        {
                            paragraph = (Paragraph)tags.get(i).getChildObjects().get(k);
                            content += paragraph.getText();
                        }
                    }
                }
            }
            property += alias + ",\t" + id + ",\t" + tag + ",\t" + sdtType + ",\t" + content + "\r\n";
            content="";
        }

        List<StructureDocumentTagRow> rowtags = structureTags.getM_rowtags();
        for (int i = 0; i < rowtags.size(); i++) {
            alias = rowtags.get(i).getSDTProperties().getAlias();
            id = rowtags.get(i).getSDTProperties().getId();
            tag = rowtags.get(i).getSDTProperties().getTag();
            sdt = rowtags.get(i).getSDTProperties().getSDTType();
            sdtType = sdt.toString();
            if (sdt == SdtType.Rich_Text || sdt == SdtType.Text)
            {
                if (rowtags.get(i).getChildObjects().getCount() > 0)
                {
                    for (int k = 0; k <  rowtags.get(i).getChildObjects().getCount(); k++) {
                        if (rowtags.get(i).getChildObjects().get(k).getDocumentObjectType() == DocumentObjectType.Paragraph)
                        {
                            paragraph = (Paragraph)rowtags.get(i).getChildObjects().get(k);
                            content += paragraph.getText();
                        }
                    }
                }
            }
            property += alias + ",\t" + id + ",\t" + tag + ",\t" + sdtType + ",\t" + content + "\r\n";
            content="";
        }
        List<StructureDocumentTagCell> celltags = structureTags.getM_celltags();
        for (int i = 0; i < celltags.size(); i++) {
            alias = celltags.get(i).getSDTProperties().getAlias();
            id = celltags.get(i).getSDTProperties().getId();
            tag = celltags.get(i).getSDTProperties().getTag();
            sdt = celltags.get(i).getSDTProperties().getSDTType();
            sdtType = sdt.toString();
            if (sdt == SdtType.Rich_Text || sdt == SdtType.Text)
            {
                if (celltags.get(i).getChildObjects().getCount() > 0)
                {
                    for (int k = 0; k <  celltags.get(i).getChildObjects().getCount(); k++) {
                        if (celltags.get(i).getChildObjects().get(k).getDocumentObjectType() == DocumentObjectType.Paragraph)
                        {
                            paragraph = (Paragraph)celltags.get(i).getChildObjects().get(k);
                            content += paragraph.getText();
                        }
                    }
                }
            }
            property += alias + ",\t" + id + ",\t" + tag + ",\t" + sdtType + ",\t" + content + "\r\n";
            content="";
        }

        //Save the property to a text document
        String output = "output/getContentControlProperty1.txt";
        File file = new File(output);
        if(!file.exists()){
            file.delete();
        }
        file.createNewFile();
        FileWriter fw = new FileWriter(file, true);
        BufferedWriter bw = new BufferedWriter(fw);
        bw.write(property);
        bw.flush();
        bw.close();
        fw.close();
    }

    public static StructureTags GetAllTags(Document document) {
        StructureTags structureTags = new StructureTags();
        for (int i = 0; i < document.getSections().getCount(); i++) {
            Section section = document.getSections().get(i);
            for (int j = 0; j < section.getBody().getChildObjects().getCount(); j++){
                DocumentObject obj = section.getBody().getChildObjects().get(j);
                if (obj.getDocumentObjectType() == DocumentObjectType.Structure_Document_Tag) {
                    structureTags.getM_tags().add((StructureDocumentTag)obj);
                }
                else if (obj.getDocumentObjectType() == DocumentObjectType.Paragraph) {
                    Paragraph para = (Paragraph)obj;
                    for (int k = 0; k <  para.getChildObjects().getCount(); k++) {
                        DocumentObject pobj = para.getChildObjects().get(k);
                        if (pobj.getDocumentObjectType() == DocumentObjectType.Structure_Document_Tag_Inline) {
                            structureTags.getM_tagInlines().add((StructureDocumentTagInline)pobj);
                        }
                    }
                }
                else if (obj.getDocumentObjectType() == DocumentObjectType.Table) {
                    Table table = (Table)obj;
                    for(int r = 0; r < table.getRows().getCount(); r++) {
                        if (table.getRows().get(r).getDocumentObjectType() == DocumentObjectType.Structure_Document_Tag_Row) {
                            structureTags.getM_rowtags().add((StructureDocumentTagRow)(table.getRows().get(r)));
                        }
                        TableRow row = table.getRows().get(r);
                        for (int c = 0; c < row.getCells().getCount(); c++) {
                            if (row.getCells().get(c).getDocumentObjectType() == DocumentObjectType.Structure_Document_Tag_Cell) {
                                structureTags.getM_celltags().add((StructureDocumentTagCell)(row.getCells().get(c)));
                            }
                            TableCell cell = row.getCells().get(c);
                            for (int s = 0; s < cell.getChildObjects().getCount(); s++) {
                                DocumentObject cellChild = cell.getChildObjects().get(s);
                                if (cellChild.getDocumentObjectType() == DocumentObjectType.Structure_Document_Tag) {
                                    structureTags.getM_tags().add((StructureDocumentTag)cellChild);
                                }
                                else if (cellChild.getDocumentObjectType() == DocumentObjectType.Paragraph) {
                                    Paragraph para = (Paragraph)cellChild;
                                    for (int t = 0; t < para.getChildObjects().getCount(); t++) {
                                        DocumentObject pobj = para.getChildObjects().get(t);
                                        if (pobj.getDocumentObjectType() == DocumentObjectType.Structure_Document_Tag_Inline) {
                                            structureTags.getM_tagInlines().add((StructureDocumentTagInline)pobj);
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
}

class StructureTags {
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

