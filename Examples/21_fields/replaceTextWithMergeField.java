import com.spire.doc.*;
import com.spire.doc.documents.*;
import com.spire.doc.fields.*;

public class replaceTextWithMergeField {
    public static void main(String[] args) {
        //Open a Word document
        Document document = new Document("data/sampleB_2.docx");

        //Find the text that will be replaced
        TextSelection ts = document.findString("Test", true, true);

        TextRange tr = ts.getAsOneRange();

        //Get the paragraph
        Paragraph par = tr.getOwnerParagraph();

        //Get the index of the text in the paragraph
        int index = par.getChildObjects().indexOf(tr);

        //Create a new field
        MergeField field = new MergeField(document);
        field.setFieldName("MergeField");

        //Insert field at specific position
        par.getChildObjects().insert(index, field);

        //Remove the text
        par.getChildObjects().remove(tr);

        //Save to file
        String output = "output/replaceTextWithMergeField.docx";
        document.saveToFile(output, FileFormat.Docx_2013);
    }
}
