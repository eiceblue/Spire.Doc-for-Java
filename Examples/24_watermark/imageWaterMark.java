import com.spire.doc.*;

public class imageWaterMark {
    public static void main(String[] args) {
        Document document = new Document("data/sampleB_2.docx");

        //Insert the imgae watermark
        insertImageWatermark(document);

        //Save as docx file
        String output = "output/imageWaterMark.docx";
        document.saveToFile(output, FileFormat.Docx);

    }

    public static void insertImageWatermark(Document document) {
        PictureWatermark picture = new PictureWatermark();
        picture.setPicture("data/imageWatermark.png");
        picture.setScaling(250);
        picture.isWashout(false);
        document.setWatermark(picture);
    }
}
