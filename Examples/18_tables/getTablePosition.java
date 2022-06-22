import com.spire.doc.*;
import com.spire.doc.formatting.TablePositioning;
import java.io.*;

public class getTablePosition {
    public static void main(String[] args) throws IOException {
        //Create a document
        Document document = new Document();
        //Load file
        document.loadFromFile("data/TableSample-Az.docx");
        //Get the first section
        Section section = document.getSections().get(0);
        //Get the first table
        Table table = section.getTables().get(0);

        StringBuilder stringBuidler = new StringBuilder();

        //Verify whether the table uses "Around" text wrapping or not.
        if (table.getTableFormat().getWrapTextAround())
        {
            TablePositioning positon = table.getTableFormat().getPositioning();

            stringBuidler.append("Horizontal:");
            stringBuidler.append("\r\n");
            stringBuidler.append("Position:" + positon.getHorizPosition() +" pt");
            stringBuidler.append("\r\n");
            stringBuidler.append("Absolute Position:" + positon.getHorizPositionAbs()+ ", Relative to:" + positon.getHorizRelationTo());
            stringBuidler.append("\r\n");

            stringBuidler.append("Vertical:");
            stringBuidler.append("\r\n");
            stringBuidler.append("Position:" + positon.getVertPosition() + " pt");
            stringBuidler.append("\r\n");
            stringBuidler.append("Absolute Position:" + positon.getVertPositionAbs() + ", Relative to:" + positon.getVertRelationTo());
            stringBuidler.append("\r\n");
            stringBuidler.append("Distance from surrounding text:");
            stringBuidler.append("\r\n");
            stringBuidler.append("Top:" + positon.getDistanceFromTop() + " pt, Left:" + positon.getDistanceFromLeft() + " pt");
            stringBuidler.append("\r\n");
            stringBuidler.append("Bottom:" + positon.getDistanceFromBottom() + "pt, Right:" + positon.getDistanceFromRight() + " pt");
        }

        //Save to txt file
        String result = "output/GetTablePosition_out.txt";
        writeStringToTxt(stringBuidler.toString(),result);
    }
    public static void writeStringToTxt(String content, String txtFileName) throws IOException {
        File file=new File(txtFileName);
        if (file.exists())
        {
            file.delete();
        }
        FileWriter fWriter = new FileWriter(txtFileName, true);
        try {
            fWriter.write(content);
        } catch (IOException ex) {
            ex.printStackTrace();
        } finally {
            try {
                fWriter.flush();
                fWriter.close();
            } catch (IOException ex) {
                ex.printStackTrace();
            }
        }
    }
}
