import java.io.IOException;
import java.net.URISyntaxException;

public class Main {
    public static void main(String[] args) throws IOException, URISyntaxException {
        var excelReader = new ExcelReader();
        var data = excelReader.read("test.xlsx");
        var excelWriter = new ExcelWriter();
        excelWriter.write("result.xlsx", data);
    }
}
