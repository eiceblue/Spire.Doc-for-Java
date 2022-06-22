import com.spire.doc.*;
import java.io.*;

public class identifyMergeFieldName {

    public static void main(String[] args) throws Exception {

        //Create Word document.
        Document document = new Document();

        //Load the file from disk.
        document.loadFromFile("data/IdentifyMergeFieldNames.docx");

        //Get the collection of group names.
        String[] GroupNames = document.getMailMerge().getMergeGroupNames();

        //Get the collection of merge field names in a specific group.
        String[] MergeFieldNamesWithinRegion = document.getMailMerge().getMergeFieldNames("Products");

        //Get the collection of all the merge field names.
        String[] MergeFieldNames = document.getMailMerge().getMergeFieldNames();

        StringBuilder content = new StringBuilder();
        content.append("----------------Group Names-----------------------------------------");
        for (int i = 0; i < GroupNames.length; i++)
        {
            content.append(GroupNames[i]);
        }

        content.append("----------------Merge field names within a specific group-----------");
        for (int j = 0; j < MergeFieldNamesWithinRegion.length; j++)
        {
            content.append(MergeFieldNamesWithinRegion[j]);
        }

        content.append("----------------All of the merge field names------------------------");
        for (int k = 0; k < MergeFieldNames.length; k++)
        {
            content.append(MergeFieldNames[k]);
        }

        String result = "output/Result-IdentifyMergeFieldNames.txt";

        writeStringToTxt(content.toString(), result);
    }
    public static void writeStringToTxt(String content, String txtFileName) throws IOException {
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
