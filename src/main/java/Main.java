import java.util.List;

public class Main {

    public static void main(String[] args) throws Exception {
        MicrosoftConnection connection = new MicrosoftConnection();
        connection.connection1();
        List<List<String>> listTableRows = connection.getListTableRows();
    }
}
