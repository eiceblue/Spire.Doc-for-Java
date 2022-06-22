import com.spire.doc.*;
import com.spire.doc.documents.BorderStyle;
import java.awt.*;
import java.io.*;

public class getDiagonalBorderProperties {
    public static void main(String[] args) throws IOException {
        String input = "output/setDiagonalBorder.docx";
        String output = "output/getDiagonalBorderProperties.txt";

        //create a document
        Document document = new Document();

        //load file
        document.loadFromFile(input);

        //get table
        Section section = document.getSections().get(0);
        Table table = section.getTables().get(0);

        //get diagonal border properties
        BorderStyle cellBorderStyle_down = table.get(0,0).getCellFormat().getBorders().getDiagonalDown().getBorderType();
        float cellBorderLineWidth_down = table.get(0,0).getCellFormat().getBorders().getDiagonalDown().getLineWidth();
        Color cellBorderColor_down = table.get(0,0).getCellFormat().getBorders().getDiagonalDown().getColor();
        BorderStyle cellBorderStyle_up = table.get(3,2).getCellFormat().getBorders().getDiagonalUp().getBorderType();
        float cellBorderLineWidth_up = table.get(3,2).getCellFormat().getBorders().getDiagonalUp().getLineWidth();
        Color cellBorderColor_up = table.get(3,2).getCellFormat().getBorders().getDiagonalUp().getColor();

        //create a new TXT file to save the text
        String result ="DiagonalUp BorderStyle" +cellBorderStyle_up+";\r\nDiagonalUp BorderColor" +cellBorderColor_up+";\r\nDiagonalUp BorderLineWidth"+cellBorderLineWidth_up+";\r\nDiagonalDown BorderStyle"+cellBorderStyle_down+";\r\n DiagonalDown BorderColor" +cellBorderColor_down+";\r\nDiagonalDown BorderLineWidth"+cellBorderLineWidth_down+";\r\n";
        writeStringToTxt(result,output);
    }
    public static void writeStringToTxt(String content, String txtFileName) throws IOException {
        FileWriter fWriter= new FileWriter(txtFileName,true);
        try {
            fWriter.write(content);
        }catch(IOException ex){
            ex.printStackTrace();
        }finally{
            try{
                fWriter.flush();
                fWriter.close();
            } catch (IOException ex) {
                ex.printStackTrace();
            }
        }
    }
}
