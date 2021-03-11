package sonar;

import org.apache.commons.csv.CSVFormat;
import org.apache.commons.csv.CSVParser;
import org.apache.commons.csv.CSVRecord;
import org.apache.poi.ss.usermodel.CellType;
import org.apache.poi.ss.util.CellReference;
import org.apache.poi.xssf.usermodel.XSSFCell;
import org.apache.poi.xssf.usermodel.XSSFRow;
import org.apache.poi.xssf.usermodel.XSSFSheet;
import org.apache.poi.xssf.usermodel.XSSFWorkbook;

import java.io.File;
import java.io.FileInputStream;
import java.io.FileNotFoundException;
import java.io.FileOutputStream;
import java.io.IOException;
import java.io.Reader;
import java.nio.file.Files;
import java.nio.file.Paths;
import java.util.HashMap;
import java.util.Map;
import java.util.Properties;

/**
 *
 */
public class UpdateXSSFWorkbook {
    private static Map<String, Integer> columnIndexMap = new HashMap<String, Integer>();
    private static Map<String, Integer> addressCellIndexMap = new HashMap<String, Integer>();
    private static Properties properties = null;

    public static void main(String[] args)
            throws IOException {

        if (args == null || args.length == 0) {
            System.err.println(" The properties file is required");
            return;
        }
        properties = load(args[0]);


        String documentQuote = properties.getProperty("sonar.quote.path");
        String documentQuoteSheet = properties.getProperty("sonar.quote.sheet");

        String sonarHeaderIndexModule = properties.getProperty("sonar.quote.header.index.module");

        String csvCountLinePath[] = properties.getProperty("csv.count.line.path").split(",");



        boolean cleanDefaultValue = Boolean.parseBoolean(properties.getProperty("sonar.quote.clean.default.value", "false"));


        // Read Quote file
        File documentQuoteFile = new File(documentQuote);
        FileInputStream inputStream = new FileInputStream(documentQuoteFile);

        // Get the workbook instance for XSSF
        XSSFWorkbook workbook = new XSSFWorkbook(inputStream);

        // Get first sheet from the workbook
        XSSFSheet sheet = workbook.getSheetAt(workbook.getSheetIndex(documentQuoteSheet));

        // File csvCountLineFile = new File(csvCountLine)
        for (int i = 0; i < csvCountLinePath.length; i++) {
            String csvFilePath = csvCountLinePath[i];
            Reader reader = Files.newBufferedReader(Paths.get(csvFilePath));
            String csvHeaderIndex = String.valueOf(i + 1);
            String[] csvHeaders = properties.getProperty("csv.count.line.headers." + csvHeaderIndex).split(",");
            String csvLineCount = properties.getProperty("csv.count.line.header.count." + csvHeaderIndex);
            String csvHeaderModule = properties.getProperty("csv.count.line.header.module." + csvHeaderIndex);
            String csvHeaderType = properties.getProperty("csv.count.line.header.language." + csvHeaderIndex);
            CSVParser parser = CSVParser.parse(reader, CSVFormat.DEFAULT
                    .withHeader(csvHeaders)
                    .withIgnoreHeaderCase()
                    .withTrim());


            for (CSVRecord csvRecord : parser) {

                if (csvRecord.getRecordNumber() == 1) {
                    continue;
                }
                String csvColumnInedexValue = csvRecord.get(csvHeaderType);

                String csvHeaderModuleValue = csvRecord.get(csvHeaderModule);

                Integer moduleRowIndex = getIndexAndRefreshMap(csvHeaderModuleValue, sheet, sonarHeaderIndexModule);


                if (moduleRowIndex == null) {
                    System.err.println("Error CSV File " + csvFilePath + " index module " + csvHeaderModuleValue + " not exit in output File");
                    continue;
                }

                //Add line code
                String columnIndex = properties.getProperty("sonar.quote.header.index.code." + csvColumnInedexValue);

                if (columnIndex != null && !"".equals(columnIndex)) {
                    boolean isInitCell = cleanDefaultValue && isFirstCellValue(columnIndex, moduleRowIndex);
                    String csvLineValue = csvRecord.get(csvLineCount);
                    genericUpdateCellByType(csvLineValue, columnIndex, sheet, moduleRowIndex, isInitCell, CellType.NUMERIC);
                } else {
                    System.err.println("Warn CSV File " + csvFilePath + " Language of " + csvColumnInedexValue + " is not specified ");
                }


                //Add template
                String columnTemplateIndex = properties.getProperty("sonar.quote.header.index.template." + csvColumnInedexValue);
                if (columnTemplateIndex != null && !"".equals(columnTemplateIndex)) {
                    boolean isInitCell = cleanDefaultValue && isFirstCellValue(columnTemplateIndex, moduleRowIndex);
                    genericUpdateCellByType("1", columnTemplateIndex, sheet, moduleRowIndex, isInitCell, CellType.NUMERIC);
                }

                //Add String
                String columnStringIndex = properties.getProperty("sonar.quote.header.index.string." + csvColumnInedexValue);
                if (columnStringIndex != null && !"".equals(columnStringIndex)) {
                    boolean isInitCell = cleanDefaultValue && isFirstCellValue(columnTemplateIndex, moduleRowIndex);
                    String csvLineValue = csvRecord.get(csvLineCount);
                    genericUpdateCellByType(csvLineValue, columnStringIndex, sheet, moduleRowIndex, isInitCell, CellType.STRING);
                }
            }
        }


        inputStream.close();

        // Write File
        FileOutputStream out = new FileOutputStream(documentQuoteFile);
        workbook.write(out);
        out.close();
    }


    /**
     * @param countHeaderModule
     * @return
     */
    public static boolean isIndexModuleExistInMap(String countHeaderModule) {
        return columnIndexMap.containsKey(countHeaderModule);
    }

    /**
     * @param countHeaderModule
     * @param sheet
     * @param sonarHeaderIndexModule
     */
    public static Integer getIndexAndRefreshMap(String countHeaderModule, XSSFSheet sheet, String sonarHeaderIndexModule) {

        if (!isIndexModuleExistInMap(countHeaderModule)) {
            int lastRowNum = sheet.getLastRowNum();
            for (int i = 0; i < lastRowNum; i++) {

                XSSFRow row = sheet.getRow(i);
                if(row != null){
                    XSSFCell cell = row.getCell(CellReference.convertColStringToIndex(sonarHeaderIndexModule));
                    if (cell != null && CellType.STRING.equals(cell.getCellTypeEnum())) {
                        if (cell.getStringCellValue().equals(countHeaderModule)) {
                            columnIndexMap.put(countHeaderModule, row.getRowNum());
                        }
                    }
                }

            }
        }
        return columnIndexMap.get(countHeaderModule);
    }


    /**
     *
     * @param cellValue
     * @param sonarHeaderIndex
     * @param sheet
     * @param moduleRowIndex
     * @param isInitCell
     * @param cellType
     */
    public static void genericUpdateCellByType(String cellValue, String sonarHeaderIndex, XSSFSheet sheet, int moduleRowIndex, boolean isInitCell, CellType cellType) {

        if (CellReference.convertColStringToIndex(sonarHeaderIndex) != -1) {

            XSSFCell cell = sheet.getRow(moduleRowIndex).getCell(CellReference.convertColStringToIndex(sonarHeaderIndex));

            if (cell == null) {
                cell = sheet.getRow(moduleRowIndex).createCell(CellReference.convertColStringToIndex(sonarHeaderIndex), cellType);
            } else if (!cell.getCellTypeEnum().equals(cellType)) {
                cell.setCellType(cellType);
            }
            if (isInitCell) {
                if (cellType.equals(CellType.NUMERIC)) {
                    cell.setCellValue(0);
                } else {
                    cell.setCellValue("");
                }

            }

            if (cellType.equals(CellType.NUMERIC)) {
                try {
                    double cellNumericValue = Double.parseDouble(cellValue);
                    System.out.println("Add Cell " + cell.getAddress() + " Somme : " + cell.getNumericCellValue() + "+" + cellValue + "=" + cell.getNumericCellValue() + cellValue + " Module Index:" + moduleRowIndex);
                    cell.setCellValue(cell.getNumericCellValue() + cellNumericValue);
                } catch (NumberFormatException e) {
                    System.err.println("headerIndex " + sonarHeaderIndex + " Value " + cellValue + "is not numeric  ");
                }

            } else {
                StringBuffer buffer = new StringBuffer();
                buffer.append(cell.getStringCellValue());
                buffer.append(cellValue.replaceAll(";", "\n"));
                System.out.println("Add Cell " + cell.getAddress() + " Value : " + buffer.toString());
                cell.setCellValue(buffer.toString());
            }


        } else {
            System.err.println(" Error with index column  " + sonarHeaderIndex);

        }

    }


    private static boolean isFirstCellValue(String columnIndex, int rowIndex) {
        boolean isFirstCellValue = false;
        String key = rowIndex + "_" + columnIndex;
        if (!addressCellIndexMap.containsKey(key)) {
            addressCellIndexMap.put(key, 0);
            isFirstCellValue = true;
        }
        return isFirstCellValue;
    }


    /**
     * @param propertiesFile
     * @return
     * @throws IOException
     * @throws FileNotFoundException
     */
    public static Properties load(String propertiesFile) throws IOException, FileNotFoundException {
        Properties properties = new Properties();

        FileInputStream input = new FileInputStream(propertiesFile);
        try {

            properties.load(input);
            return properties;

        } finally {

            input.close();

        }

    }

}
