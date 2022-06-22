import com.spire.doc.*;
import com.spire.doc.documents.*;

public class setImageBackground {
    public static void main(String[] args) {
        //Load a word document
        Document document = new Document("data/Template.docx");

        //Set the background type as picture.
        document.getBackground().setType(BackgroundType.Picture);

        //Set the background picture
        document.getBackground().setPicture("data/Background.png");

        String result = "output/result-ImageBackground.docx";
        //Save the file.
        document.saveToFile(result, FileFormat.Docx);

    }
}
