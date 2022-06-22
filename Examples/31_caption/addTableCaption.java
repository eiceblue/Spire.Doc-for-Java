import com.spire.doc.*;

public class addTableCaption {
    public static void main(String[] args) {
        //Create word document
        Document document = new Document();
        //Load file
        document.loadFromFile("data/tableTemplate.docx");

        //Get the first table
        Body body = document.getSections().get(0).getBody();
        Table table = (Table)body.getTables().get(0);

        //Add caption to the table
        table.addCaption("Table", CaptionNumberingFormat.Number, CaptionPosition.Below_Item);

        //Update fields
        document.isUpdateFields(true);

        //Save the file
        String output = "output/addTableCaption_result.docx";
        document.saveToFile(output, FileFormat.Docx);
    }
}
