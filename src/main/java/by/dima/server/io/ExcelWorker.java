package by.dima.server.io;

import lombok.SneakyThrows;
import lombok.extern.slf4j.Slf4j;
import org.apache.poi.ss.usermodel.Cell;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.usermodel.CellValue;
import org.apache.poi.ss.usermodel.FormulaEvaluator;
import org.apache.poi.ss.usermodel.Row;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;
import org.springframework.beans.factory.annotation.Value;
import org.springframework.stereotype.Component;

import java.io.FileInputStream;
import java.io.InputStream;

@Component
@Slf4j
public class ExcelWorker {

    @Value(value = "${file.name}")
    private String fileName;

    @SneakyThrows
    public String read() {

        try (InputStream is = new FileInputStream(fileName);
             XSSFWorkbook wb = new XSSFWorkbook(is)) {
            FormulaEvaluator evaluator = wb.getCreationHelper().createFormulaEvaluator();
            XSSFSheet sheet = wb.getSheetAt(0);
            for (int i = 0; i < sheet.getLastRowNum(); i++) {
                Row row = sheet.getRow(i);
                for (int cn = 0; cn < row.getLastCellNum(); cn++) {
                    Cell cell = row.getCell(cn, Row.MissingCellPolicy.RETURN_BLANK_AS_NULL);

                    if (cell == null) {
                        System.out.println("Ячейка пуста (" + cn + ")");
                    } else {
                        if (cell.getCellType() == CellType.FORMULA){
                            CellValue value = evaluator.evaluate(cell);

                            switch (value.getCellType()) {
                                case STRING:
                                    System.out.println("Строка: " + value.getStringValue());
                                    break;
                                case NUMERIC:
                                    System.out.println("Число: " + value.getNumberValue());
                                    break;
                                case BOOLEAN:
                                    System.out.println("Булево: " + value.getBooleanValue());
                                    break;
                                case ERROR:
                                    System.out.println("Ошибка: " + value.getErrorValue());
                                    break;
                            }
                        }
                    }
                }
            }
        }

        return new String();
    }

}
