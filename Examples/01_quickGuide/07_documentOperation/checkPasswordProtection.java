import com.spire.doc.*;

public class checkPasswordProtection {
    public static void main(String[] args) {
       boolean value = Document.isPassWordProtected("data/CheckPasswordProtection.docx");
       if(value) {
           System.out.println("This Word document is password protected.");
       }else {
           System.out.println("This Word document is not password protected.");
       }
    }
}
