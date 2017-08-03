import java.io.IOException;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.Paths;
import java.util.HashMap;
import java.util.List;
import java.util.Map;

/* 

Поиск текста внутри xls-файлов

*/
public class Solution {

    public static void main(String[] args) throws IOException {
        SearchFileVisitor searchFileVisitor = new SearchFileVisitor();

        //searchFileVisitor.setPartOfName("amigo");
        searchFileVisitor.setPartOfContent("розетка");
        searchFileVisitor.setMatchCase(false);
        //searchFileVisitor.setMinSize(500);
        //searchFileVisitor.setMaxSize(10000);

        Files.walkFileTree(Paths.get("FileSearch"), searchFileVisitor);

        HashMap<Path, List<String>> foundFiles = searchFileVisitor.getFoundFiles();

        for(Map.Entry<Path, List<String>> set : foundFiles.entrySet()){
            System.out.printf("имя файла: \"%s\":", set.getKey().getFileName());
            System.out.println();
            for (String s: set.getValue()                 ) {
                System.out.printf("\t%s%s", s, System.lineSeparator());
            }
            System.out.println();
        }

    }

}
