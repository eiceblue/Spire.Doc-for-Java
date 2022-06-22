import com.spire.doc.*;


public class setEditableRange {
    public static void main(String[] args) {
        String inputFile = "data/setEditableRange.docx";
        String outputFile = "setEditableRange_output.docx";
        String password = "E-iceblue";
        String permissionId = "myId";
        //Create and load the Word file from the disk.
        Document document = new Document();
        document.loadFromFile(inputFile);
        //Get first section
        Section section = document.getSections().get(0);
        //Create start and end tags.
        PermissionStart start = new PermissionStart(document, permissionId);
        PermissionEnd end = new PermissionEnd(document, permissionId);

        //Set editable range by inserting the start and end tags.
        section.getParagraphs().get(0).getChildObjects().insert(0, start);
        section.getParagraphs().get(0).getChildObjects().add(end);
        //After protecting, all content will not be able to edit except the first paragraph of the first section.
        document.protect(ProtectionType.Allow_Only_Reading, password);
        //Save the document
        document.saveToFile(outputFile, FileFormat.Docx);
    }
}
