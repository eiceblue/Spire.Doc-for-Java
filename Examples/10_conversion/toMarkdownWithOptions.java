import com.spire.doc.*;

public class toMarkdownWithOptions {
    public static void main(String[] args) {
        // Create a new Document instance.
        Document document = new Document();

        // Load the content of an existing .docx file.
        document.loadFromFile("Data/toMarkdown.docx");

        // Configure the Markdown export options to embed images as Base64-encoded strings instead of saving them as separate files.
        //document.getMarkdownExportOptions().setImagesAsBase64(true);

        // Specify the local directory where extracted images from the document will be saved during Markdown conversion.
        document.getMarkdownExportOptions().setImagesFolder("D:\\Markdown\\Images");

        // Set an alias path for the images folder that will be used in the generated Markdown file.
        // document.getMarkdownExportOptions().setImagesFolderAlias("D:\\Markdown\\MyImages");

        // Set the list output mode to render bullet and numbered lists as plain text without Markdown list syntax.
        document.getMarkdownExportOptions().setListOutputMode(MarkdownListOutputMode.Plain_Text);

        // Enable preservation of underline formatting in the exported Markdown content.
        document.getMarkdownExportOptions().setSaveUnderlineFormatting(true);

        // Configure hyperlink output to use Markdown reference-style links (e.g., [text][1] with definitions at the bottom).
        document.getMarkdownExportOptions().setLinkOutputMode(MarkdownLinkOutputMode.Reference);

        // Set the output format for Office Math equations to MathML within the Markdown file.
        document.getMarkdownExportOptions().setOfficeMathOutputMode(MarkdownOfficeMathOutputMode.Math_ML);

        // Specify that certain elements (e.g., tables) should be saved using HTML syntax within the Markdown output.
        document.getMarkdownExportOptions().setSaveAsHtml(MarkdownSaveAsHtml.Tables);

        // Save the converted document to a Markdown file.
        document.saveToFile("toMarkdown_out.md", FileFormat.Markdown);

        // Close the document to release associated resources.
        document.close();

        // Explicitly dispose of the document object to free memory and system resources.
        document.dispose();
    }
}
