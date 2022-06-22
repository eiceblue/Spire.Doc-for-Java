import com.spire.doc.*;
public class getDocumentProperties {
    public static void main(String[] args) {
        //Create a document
        Document document = new Document();
        //Load the document from disk.
        document.loadFromFile("data/Properties.docx");

        //Create StringBuilder to save
        StringBuilder content = new StringBuilder();

        //Get Builtin document properties
        String title = document.getBuiltinDocumentProperties().getTitle();
        String comments = document.getBuiltinDocumentProperties().getComments();
        String author = document.getBuiltinDocumentProperties().getAuthor();
        String keywords = document.getBuiltinDocumentProperties().getKeywords();
        String company = document.getBuiltinDocumentProperties().getCompany();

        //Set string format for displaying
        String result = String.format("The Builtin document properties:\r\nTitle: " + title + ".\r\nComments: " + comments + ".\r\nAuthor: " + author + ".\r\nKeywords: " + keywords + ".\r\nCompany: " + company);

        //Add result string to StringBuilder
        content.append(result + "\r\nThe custom document properties:");

        //Get custom document properties
        for (int i = 0; i < document.getCustomDocumentProperties().getCount(); i++) {
            String customProperties = String.format(document.getCustomDocumentProperties().get(i).getName() + ": " + document.getCustomDocumentProperties().get(i).getValue());
            content.append(customProperties);
        }
        System.out.println(content);
    }
}
