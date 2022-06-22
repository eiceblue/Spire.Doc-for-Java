import com.spire.doc.*;
import com.spire.doc.documents.*;
import com.spire.doc.fields.*;
import javax.imageio.ImageIO;
import java.awt.image.*;
import java.io.*;

public class convertObjectToImage {
    public static void main(String[] args) throws Exception {
        //Create a document
        Document document = new Document();
        //Load file
        document.loadFromFile("data/ConvertObjectToImage.docx");
        //Get the first section
        Section section = document.getSections().get(0);
        //Get body of section
        Body body = section.getBody();

        //Get the first paragraph
        Paragraph paragraph = body.getParagraphs().get(0);
        BufferedImage image = ConvertParagraphToImage(paragraph);
        File file = new File("output/convertParagraphToImage.png");
        ImageIO.write(image, "png", file);

        //Get the first table
        Table table = (Table) body.getTables().get(0);
        image = ConvertTableToImage(table);
        file = new File("output/ConvertTableToImage.png");
        ImageIO.write(image, "png", file);

        //Get the first row of the first table
        TableRow row = table.getRows().get(0);
        image = ConvertTableRowToImage(row);
        file = new File("output/convertTableRowToImage.png");
        ImageIO.write(image, "png", file);

        //Get the first cell of the first row
        TableCell cell = row.getCells().get(0);
        image = ConvertTableCellToImage(cell);
        file = new File("output/convertTableCellToImage.png");   //Get a shape
        int i = 0;
        for (int j = 0; j < section.getParagraphs().getCount(); j++) {
            Paragraph p = section.getParagraphs().get(j);
            for (int k = 0; k < p.getChildObjects().getCount(); k++) {
                DocumentObject obj = p.getChildObjects().get(k);
                if (obj.getDocumentObjectType().equals(DocumentObjectType.Shape)) {
                    image = ConvertShapeToImage((ShapeObject) obj);
                    file = new File("output/convertShapeToImage-" + i + ".png");
                    ImageIO.write(image, "png", file);
                    i++;
                }
            }
        }
        ImageIO.write(image, "png", file);
    }

    private static BufferedImage ConvertParagraphToImage(Paragraph obj) {
        Document doc = new Document();
        Section section = doc.addSection();
        section.getBody().getChildObjects().add(obj.deepClone());
        BufferedImage image = doc.saveToImages(0, ImageType.Bitmap);
        doc.close();
        return image;
    }

    private static BufferedImage ConvertTableToImage(Table obj) {
        Document doc = new Document();
        Section section = doc.addSection();
        section.getBody().getChildObjects().add(obj.deepClone());
        BufferedImage image = doc.saveToImages(0, ImageType.Bitmap);
        doc.close();
        return image;
    }

    private static BufferedImage ConvertTableRowToImage(TableRow obj) {
        Document doc = new Document();
        Section section = doc.addSection();
        Table table = section.addTable();
        table.getRows().add(obj.deepClone());
        BufferedImage image = doc.saveToImages(0, ImageType.Bitmap);
        doc.close();
        return image;
    }

    private static BufferedImage ConvertTableCellToImage(TableCell obj) {
        Document doc = new Document();
        Section section = doc.addSection();
        Table table = section.addTable();
        table.addRow().getCells().add(obj.deepClone());
        BufferedImage image = doc.saveToImages(0, ImageType.Bitmap);
        doc.close();
        return image;
    }

    private static BufferedImage ConvertShapeToImage(ShapeObject obj) {
        Document doc = new Document();
        Section section = doc.addSection();
        section.addParagraph().getChildObjects().add(obj.deepClone());
        BufferedImage image = doc.saveToImages(0, ImageType.Bitmap);
        doc.close();
        return image;
    }
}
