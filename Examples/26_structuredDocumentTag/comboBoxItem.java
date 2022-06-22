import com.spire.doc.*;
import com.spire.doc.documents.*;

public class comboBoxItem {
    public static void main(String[] args) {
        // Create a new document and load from file
        Document doc = new Document();
        doc.loadFromFile("data/comboBox.docx");

        //Get the combo box from the file
        for (int i = 0; i < doc.getSections().getCount(); i++) {
            Section section = doc.getSections().get(i);
            for (int j = 0; j < section.getBody().getChildObjects().getCount(); j++) {
                DocumentObject bodyObj = section.getBody().getChildObjects().get(j);
                if (bodyObj.getDocumentObjectType() == DocumentObjectType.Structure_Document_Tag) {
                    StructureDocumentTag sdt = (StructureDocumentTag) bodyObj;
                    //If SDTType is ComboBox
                    if ((sdt.getSDTProperties().getSDTType() == SdtType.Combo_Box)) {
                        SdtComboBox combo = (SdtComboBox) sdt.getSDTProperties().getControlProperties();
                        //Remove the second list item
                        combo.getListItems().removeAt(1);
                        //Add a new item
                        SdtListItem item = new SdtListItem("D", "D");
                        combo.getListItems().add(item);

                        //If the value of list items is "D"
                        for (int k = 0; k < combo.getListItems().getCount(); k++) {
                            SdtListItem sdtItem = combo.getListItems().get(k);
                            if ("D".equals(sdtItem.getValue())) {
                                //Select the item
                                combo.getListItems().setSelectedValue(sdtItem);
                            }
                        }
                    }
                }
            }
        }

        //Save the document
        String output = "output/comboBoxItem.docx";
        doc.saveToFile(output, FileFormat.Docx_2013);
    }
}
