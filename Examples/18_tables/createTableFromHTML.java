import com.spire.doc.*;

public class createTableFromHTML {
    public static void main(String[] args) {
        //HTML string
        String HTML = "<table border='2px'>" +
                "<tr>" +
                "<td>Row 1, Cell 1</td>" +
                "<td>Row 1, Cell 2</td>" +
                "</tr>" +
                "<tr>" +
                "<td>Row 2, Cell 2</td>" +
                "<td>Row 2, Cell 2</td>" +
                "</tr>" +
                "</table>";

        //Create a Word document
        Document document = new Document();

        //Add a section
        Section section = document.addSection();

        //Add a paragraph and append html string
        section.addParagraph().appendHTML(HTML);

        //Save to Word document
        String output = "output/CreateTableFromHTML_out.docx";
        document.saveToFile(output, FileFormat.Docx_2013);
    }
}
