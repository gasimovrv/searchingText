import com.sun.star.comp.helper.ComponentContext;
import com.sun.star.frame.XComponentLoader;
import com.sun.star.frame.XDesktop;
import com.sun.star.sheet.XSpreadsheetDocument;
import com.sun.star.uno.UnoRuntime;
import com.sun.star.uno.XComponentContext;
import org.apache.poi.hssf.usermodel.HSSFWorkbook;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.ss.usermodel.Sheet;
import com.sun.star.frame.XDesktop;

import java.io.FileInputStream;
import java.io.IOException;
import java.nio.file.FileVisitResult;
import java.nio.file.Files;
import java.nio.file.Path;
import java.nio.file.SimpleFileVisitor;
import java.nio.file.attribute.BasicFileAttributes;
import java.util.ArrayList;
import java.util.HashMap;
import java.util.Hashtable;
import java.util.List;


/**
* Класс для прохождения по дереву файлов
**/
public class SearchFileVisitor extends SimpleFileVisitor<Path> {
    //список найденных файлов (key) и список строк (value), в которых найден искомый текст
    private HashMap<Path, List<String>> foundFiles = new HashMap<>();
    //список строк, в которых найден искомый текст
    private List<String> foundLines = new ArrayList<>();
    //условия поиска
    private String partOfName = null;
    private String partOfContent = null;
    private int minSize = -1;
    private int maxSize = -1;
    private boolean matchCase = true;

    //состояние условий поиска(true - задано, false - пропущено)
    private boolean condition1 = false;
    private boolean condition2 = false;
    private boolean condition3 = false;
    private boolean condition4 = false;


    public HashMap<Path, List<String>> getFoundFiles() {
        return foundFiles;
    }

    public void setPartOfName(String partOfName) {
        this.partOfName = partOfName;
    }

    public void setPartOfContent(String partOfContent) {
        if (matchCase) this.partOfContent = partOfContent;
        else this.partOfContent = partOfContent.toLowerCase();
    }

    public void setMinSize(int minSize) {
        this.minSize = minSize;
    }

    public void setMaxSize(int maxSize) {
        this.maxSize = maxSize;
    }

    public void setMatchCase(boolean matchCase) {
        this.matchCase = matchCase;
        if (!matchCase) this.partOfContent = partOfContent.toLowerCase();
    }

    @Override
    public FileVisitResult visitFile(Path file, BasicFileAttributes attrs) throws IOException {

        if (partOfName != null) {//если услоовие1 задано, то проверяем на соответствие
            if (file.getFileName().toString().contains(partOfName)) {
                condition1 = true;
            }
        } else //иначе принимаем условие1 выполненым
            condition1 = true;

        if (partOfContent != null) {//если услоовие2 задано, то проверяем на соответствие
            try {

                if (file.getFileName().toString().endsWith(".xls")) {//если табличный файл(ексель)
                    foundLines = searchFromExcel(file, partOfContent, matchCase);
                    if (foundLines != null)
                        condition2 = true;
                } else if (file.getFileName().endsWith(".ods")) {//если табличный файл(опен офис)
                    foundLines = searchFromOpenOffice(file, partOfContent, matchCase);
                    if (foundLines != null)
                        condition2 = true;
                } else { //другой текстовый файл
                    foundLines = Files.readAllLines(file);
                    for (String s : foundLines) {
                        if (!matchCase) {
                            s = s.toLowerCase();
                        }
                        if (s.contains(partOfContent)) {
                            condition2 = true;
                        }
                    }
                }
            } catch (IOException e) {
                e.printStackTrace();
                System.out.println("Ошибка чтения файла: " + e);
            }

        } else //иначе принимаем условие2 выполненым
            condition2 = true;

        if (minSize != -1) {//если услоовие3 задано, то проверяем на соответствие
            if (Files.size(file) >= minSize) {
                condition3 = true;
            }
        } else //иначе принимаем условие3 выполненым
            condition3 = true;

        if (maxSize != -1) {//если услоовие4 задано, то проверяем на соответствие
            if (Files.size(file) <= maxSize) {
                condition4 = true;
            }
        } else //иначе принимаем условие4 выполненым
            condition4 = true;


        if (condition1 && condition2 && condition3 && condition4) {
            foundFiles.put(file, foundLines);
        }
        condition1 = condition2 = condition3 = condition4 = false;

        return FileVisitResult.CONTINUE;
    }

    /**
     * Поиск текста внутри файла .xls
     *
     * @param file       файл с расширением .xls
     * @param targetText текст для поиска
     * @param matchCase  если true - то учитывается регистр, иначе - не учитывается
     * @return true - если хоть одно совпадение найдено
     */
    public static ArrayList<String> searchFromExcel(Path file, String targetText, boolean matchCase) throws IOException {
        ArrayList<String> result = new ArrayList<>();

        // формируем из файла экземпляр HSSFWorkbook
        HSSFWorkbook book = new HSSFWorkbook(new FileInputStream(file.toFile()));

        // получаем Iterator по всем листам
        for (Sheet sheet : book) {
            // получаем Iterator по всем строкам в листе
            for (Row row : sheet) {
                // получаем Iterator по всем ячейкам в строке
                for (Cell cell : row) {

                    if (cell.getCellTypeEnum().equals(CellType.STRING)) {
                        String cellStr = cell.getStringCellValue();
                        if (!matchCase) {
                            cellStr = cellStr.toLowerCase();
                        }
                        if (cellStr.contains(targetText)) {
                            result.add(String.format("строка №%d - \"%s\"", row.getRowNum() + 1, cell.getStringCellValue()));
                        }
                    }
                }
            }
        }
        book.close();
        return result.size() > 0 ? result : null;
    }

    public static ArrayList<String> searchFromOpenOffice(Path file, String targetText, boolean matchCase) throws IOException {
        ArrayList<String> result = new ArrayList<>();
//        XSpreadsheetDocument xssd = UnoRuntime.queryInterface(XSpreadsheetDocument.class, xSpreadsheetComponent);
        return result;
    }
}
