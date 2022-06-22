import com.spire.doc.*;
import jdk.nashorn.internal.runtime.regexp.joni.Regex;

import java.util.regex.Pattern;

public class replaceTextInTable {
    public static void main(String[] args) {
        //Load Word from disk
        Document doc = new Document();
        doc.loadFromFile("data/ReplaceTextInTable.docx");

        //Get the first section
        Section section = doc.getSections().get(0);

        //Get the first table in the section
        Table table = section.getTables().get(0);

        //Define a regular expression to match the {} with its content
        Pattern regex =  Pattern.compile("\\{[^\\}]+\\}",0);

        //Replace the text of table with regex
        table.replace(regex, "E-iceblue");

        //Replace old text with new text in table
        table.replace("Beijing", "Component", false, true);

        //Save the Word document
        String output="output/ReplaceTextInTable_out.docx";
        doc.saveToFile(output, FileFormat.Docx_2013);
    }
}
