import com.spire.doc.*;
import com.spire.doc.documents.*;
import com.spire.doc.fields.Field;

public class resetPageNumber {
    public static void main(String[] args) {
        //Create three Word documents and load three different Word documents from disk.
        Document document1 = new Document();
        document1.loadFromFile("data/ResetPageNumber1.docx");

        Document document2 = new Document();
        document2.loadFromFile("data/ResetPageNumber2.docx");

        Document document3 = new Document();
        document3.loadFromFile("data/ResetPageNumber3.docx");

        //Use section method to combine all documents into one word document.
        for (int i = 0; i < document2.getSections().getCount(); i++) {
            document1.getSections().add(document2.getSections().get(i).deepClone());
        }
        for (int i = 0; i < document3.getSections().getCount(); i++) {
            document1.getSections().add(document3.getSections().get(i).deepClone());
        }

        //Traverse every section of document1.
        for (int i = 0; i < document1.getSections().getCount(); i++) {
            Section sec = document1.getSections().get(i);
            //Traverse every object of the footer.
            for (int j = 0; j < sec.getHeadersFooters().getFooter().getChildObjects().getCount(); j++) {
                DocumentObject obj = sec.getHeadersFooters().getFooter().getChildObjects().get(j);
                if (obj.getDocumentObjectType().equals(DocumentObjectType.Structure_Document_Tag)) {
                    DocumentObject para = obj.getChildObjects().get(0);
                    for (int k = 0; k < para.getChildObjects().getCount(); k++) {
                        DocumentObject item = para.getChildObjects().get(k);
                        if (item.getDocumentObjectType().equals(DocumentObjectType.Field))
                            //Find the item and its field type is FieldNumPages.
                            if (((Field) item).getType().equals(FieldType.Field_Num_Pages)) {
                                //Change field type to FieldSectionPages.
                                ((Field) item).setType(FieldType.Field_Section_Pages);
                            }
                    }
                }
            }
        }

        //Restart page number of section and set the starting page number to 1.
        document1.getSections().get(1).getPageSetup().setRestartPageNumbering(true);
        document1.getSections().get(1).getPageSetup().setPageStartingNumber(1);

        document1.getSections().get(2).getPageSetup().setRestartPageNumbering(true);
        document1.getSections().get(2).getPageSetup().setPageStartingNumber(1);

        String result = "output/result-resetPageNumber.docx";

        //Save to file.
        document1.saveToFile(result, FileFormat.Docx_2013);
    }
}
