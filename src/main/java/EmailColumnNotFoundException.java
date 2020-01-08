public class EmailColumnNotFoundException extends Exception {
    public EmailColumnNotFoundException(){
        super("No column with name 'E-Mail' found. Please check Excel Sheet");
    }
}
