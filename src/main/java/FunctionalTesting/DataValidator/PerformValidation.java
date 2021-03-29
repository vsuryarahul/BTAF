package FunctionalTesting.DataValidator;

import FunctionalTesting.DataModel.*;
import FunctionalTesting.ExtractData.ExtractTable;
import FunctionalTesting.ExtractData.SheetHandler;
import FunctionalTesting.TestReport.ValidationTestReport;
import com.aventstack.extentreports.ExtentTest;
import com.aventstack.extentreports.Status;
import org.apache.poi.ss.usermodel.*;
import tech.tablesaw.api.Table;
import tech.tablesaw.columns.Column;

import java.io.File;
import java.io.IOException;
import java.time.Duration;
import java.time.LocalDateTime;
import java.time.format.DateTimeFormatter;
import java.util.*;

public class PerformValidation {
    ExtractTable extractTable;
    ValidationTestReport generateReport;
    String userName = System.getProperty("user.name");
    String filePath = "C:\\Users\\"+ userName +"\\Documents\\BTAF Framework";
    List<InputFormValidationData> inputFormValDataList;
    List<BackendReportValidationData> backendDValDataList;
    List<ValidationDataResults> valDataResultList;
    List<String> inputFormIntersecName;
    InputFormValidationData val;
    BackendReportValidationData bval;
    ValidationDataResults valResults;
    boolean isInputFormValidation;
    boolean isBackendReportValidation;


    //Validates the data from a report to the corresponding input form
    public String validateData(InputWbData inputWbData, ExtentTest test) throws Exception {
            List<DependencyWbData> dependencyDataList = inputWbData.getDependencyWbData();
            DateTimeFormatter dtf = DateTimeFormatter.ofPattern("HH:mm:ss");
            LocalDateTime startTime = LocalDateTime.now();
            inputWbData.getTestCaseDetails().setExecutionStartTime(dtf.format(startTime));
            String index = inputWbData.getTestCaseDetails().getMasterIndex();

            File report = inputWbData.getReportWb();
            String reportName = inputWbData.getTestCaseDetails().getReportWbName();
            String reportSheetName = inputWbData.getTestCaseDetails().getReportWsName();
            extractTable = new ExtractTable();
            inputFormValDataList = new ArrayList<>();
            backendDValDataList = new ArrayList<>();
            valDataResultList = new ArrayList<>();

            Map<String, Table> inputFormDIMTable = new HashMap<>();
            Map<String, Map<String, List<Object>>> inputFormTables = new HashMap<>();
            List<String> dependencyWbNames = new ArrayList<>();
            dependencyDataList.forEach(e -> {
                File inputForm = e.getDependencyWb();
                List<String> inputWorksheets = e.getDependencyWsNamesList();
                String fileName = inputForm.getName().contains("(") ? inputForm.getName().split("\\(")[0].trim() : inputForm.getName().split("\\.")[0].trim();
                dependencyWbNames.add(fileName);
                inputWorksheets.forEach(worksheet -> {
                    try {
                        inputFormDIMTable.put(fileName + worksheet, extractTable.convertExcelToTable(filePath + "\\Data Intersection Models\\" + fileName + ";" + worksheet + ".xlsx", worksheet + "_DIM"));
                        Table inputFileDef = extractTable.convertExcelToTable(filePath + "\\Data Intersection Models\\" + fileName + ";" + worksheet + ".xlsx", worksheet + "_TableDef");
                        inputFormTables.put(fileName + worksheet, createTablesFromUserInputs(inputFileDef, inputForm, worksheet));
                    } catch (IOException ioException) {
                        ioException.printStackTrace();
                    }
                });
            });
            inputWbData.setDependencyWbData(dependencyDataList);
            Table mappingTable = extractTable.convertExcelToTable(filePath + "\\Report Mappings\\" + index + ".xlsx", index);
            test.log(Status.INFO,mappingTable.getString(0,1));
            mappingTable.forEach(e-> {
                for(int i=0; i< e.columnCount(); i++) {
                    test.log(Status.INFO,mappingTable.getString(e.getRowNumber(),i));
                }
            });
            Table reportDIMTable = extractTable.convertExcelToTable(filePath + "\\Data Intersection Models\\" + reportName + ";" + reportSheetName.trim() + ".xlsx", reportSheetName.trim() + "_DIM");
            Table reportTableDef = extractTable.convertExcelToTable(filePath + "\\Data Intersection Models\\" + reportName + ";" + reportSheetName.trim() + ".xlsx", reportSheetName.trim() + "_TableDef");
            Map<String, List<Object>> reportTables = createTablesFromUserInputs(reportTableDef, report, reportSheetName);
            for (int i = 0; i < mappingTable.rowCount(); i++) {
                isInputFormValidation = false;
                isBackendReportValidation = false;
                boolean isReportMappingAvailable = true;
                boolean isInputFormMappingAvailable = true;
                inputFormIntersecName = new ArrayList<>();
                String dataType = mappingTable.getString(i,"Type");
                valResults = new ValidationDataResults();
                val = new InputFormValidationData();
                bval = new BackendReportValidationData();
                if(dataType.equals("Input Form")) {
                    isInputFormValidation = true;
                } else if(dataType.equals("Backend")) {
                    isBackendReportValidation = true;
                }
                String reportIntersectionName = mappingTable.getString(i, "Report Intersection");
                String inputIntersectionName = mappingTable.getString(i, "Input Form Intersection");
                if (reportIntersectionName.equals("") || inputIntersectionName.equals("")) {
                    continue;
                }
                else {
                    List<Object> reportResult= evaluate("Report Intersection", reportIntersectionName, inputIntersectionName, reportTables, reportDIMTable,
                            reportSheetName, inputFormTables, inputFormDIMTable, dependencyDataList, val, isInputFormValidation, isBackendReportValidation, valResults, bval);
                        isReportMappingAvailable = (boolean) reportResult.get(1);
                        assignValue("Report Intersection",reportIntersectionName, (Double) reportResult.get(0), valResults, val, bval);

                    List<Object> inputResult = evaluate("Input Form Intersection",reportIntersectionName, inputIntersectionName,reportTables, reportDIMTable, reportSheetName,
                            inputFormTables, inputFormDIMTable, dependencyDataList, val,isInputFormValidation, isBackendReportValidation, valResults, bval);
                    isInputFormMappingAvailable = (boolean) inputResult.get(2);
                    assignValue("Input Form Intersection",inputIntersectionName, (Double) inputResult.get(0), valResults, val, bval);


                    if(isReportMappingAvailable == false && isInputFormMappingAvailable == false) {
                        continue;
                    } else if((isReportMappingAvailable == false && isInputFormMappingAvailable == true) ||  (isReportMappingAvailable == true && isInputFormMappingAvailable == false)) {
                        valResults.setStatus("Error");
                        if(isInputFormValidation) {
                            val.setStatus("Error");
                        } else if(isBackendReportValidation) {
                            bval.setStatus("Error");
                        }
                    } else {
                        if(valResults.getReportData() == valResults.getInputFormData()) {
                            valResults.setStatus("Pass");
                            if(isInputFormValidation) {
                                val.setStatus("Pass");
                            } else if(isBackendReportValidation) {
                                bval.setStatus("Pass");
                            }
                        }
                        else {
                            valResults.setStatus("Fail");
                            if(isInputFormValidation) {
                                val.setStatus("Fail");
                            } else if(isBackendReportValidation) {
                                bval.setStatus("Fail");
                            }
                        }
                    }
                }
                if(isInputFormValidation) {
                    inputFormValDataList.add(val);
                } else if(isBackendReportValidation) {
                    backendDValDataList.add(bval);
                }
                valDataResultList.add(valResults);
            }
            LocalDateTime endTime = LocalDateTime.now();
            inputWbData.getTestCaseDetails().setExecutionEndTime(dtf.format(endTime));
            String elapsedTime = Duration.between(endTime, startTime).toMinutes() > 0 ? Duration.between(startTime, endTime).toMinutes() + " minutes" :
                    Duration.between(startTime, endTime).getSeconds() + " seconds";
            inputWbData.getTestCaseDetails().setElapsedTime(elapsedTime);
            generateReport = new ValidationTestReport();
            try {
                generateReport.generateTestReport(inputWbData.getTestCaseDetails(), inputFormValDataList, backendDValDataList, valDataResultList);
                return "Validation Completed Successfully.";
            } catch (IOException e) {
                e.printStackTrace();
            }
        return null;
    }

    //Creates a table from file uploaded by the user
    public Map<String, List<Object>> createTablesFromUserInputs(Table tableDef, File userInputFile, String userInputSheetName) throws IOException {
        List<Object> tableData = new ArrayList<>();
        Workbook wb = WorkbookFactory.create(userInputFile);
        Sheet sheet = SheetHandler.getSheetFromWorkBook(wb, 0);
        boolean isNumeric = false;
        Map<String, Map<String, List<String>>> entireTableDefData = extractTable.getTableDefinitions(tableDef);
        Table table = null;
        Map<String, List<Object>> userInputTables = new HashMap<>();

        for (Map.Entry<String, Map<String, List<String>>> e : entireTableDefData.entrySet()) {
            List<String> columnWithPercentage = new ArrayList<>();

            String tableName = e.getKey();
            Map<String, List<String>> tableDefData = e.getValue();

            for(int i=0; i<=sheet.getLastRowNum();) {
                Row row = sheet.getRow(i);
                for(int j=0;j<=row.getLastCellNum();j++) {
                    Cell cell = row.getCell(j);
                    if(cell != null)
                        if(cell != null && cell.getCellType().equals(CellType.NUMERIC)) {
                            isNumeric = true;
                            if(cell.getCellStyle().getDataFormatString().contains("%")) {
                                columnWithPercentage.add(sheet.getRow(i-1).getCell(j).getStringCellValue());
                            }
                        }
                }
                if(!isNumeric) {
                    i++;
                }
            }
            tableData.add(columnWithPercentage);

            try {
                table = extractTable.convertExcelToTableWithCellRange(userInputFile, userInputSheetName, tableDefData);
            } catch (Exception exception) {
                exception.printStackTrace();
            }
            tableData.add(table);
            userInputTables.put(tableName,tableData);
        }

        return userInputTables;
    }

    //Sets the values from the report and input form for the HTML report and labels them as either an input form validation or backend report validation
    private long assignValue(String source, String intersectionName, Double inputFileValue, ValidationDataResults valResults, InputFormValidationData val, BackendReportValidationData bval) {
        long value;
        /*if (intersectionName.contains("ACR YoY%") || intersectionName.contains("ACR Comparison Report _YOY")) {
            value = (int) Math.round(inputFileValue * 100);
        } else {
            value = Math.round(inputFileValue);
        }*/
        value = Math.round(inputFileValue);
        if(source.equals("Input Form Intersection")) {
            valResults.setInputFormData(value);
            if(isInputFormValidation) {
                    val.setInputFormData(value);
            } else if(isBackendReportValidation) {
                    bval.setInputFormData(value);
            }
        } else {
            valResults.setReportData(value);
            if(isInputFormValidation) {
                val.setReportData(value);
            } else if(isBackendReportValidation) {
                bval.setReportData(value);
            }
        }
        return value;
    }

    //Evaluates the values based on the Report intersection mapping
    public static List<Object> evaluate(String source, String reportIntersectionName, String inputIntersectionName, Map<String, List<Object>> reportTables, Table reportDIMTable, String reportSheetName, Map<String, Map<String, List<Object>>> inputFormTables, Map<String, Table> inputFormDIMTable,
                                        List<DependencyWbData> dependencyDataList, InputFormValidationData val, boolean isInputFormValidation, boolean isBackendReportValidation, ValidationDataResults valResults, BackendReportValidationData bval)
    {
        List<Object> resultList = new ArrayList<>();
        boolean isReportMappingAvailable = true;
        boolean isInputFormMappingAvailable = true;
        char[] tokens =  source.equals("Report Intersection") ? reportIntersectionName.toCharArray() : inputIntersectionName.toCharArray();
        Stack<Double> values = new Stack<>();
        Stack<Character> ops = new Stack<>();
        for (int i = 0; i < tokens.length; i++)
        {
            if (tokens[i] == ' ')
                continue;
            if (tokens[i] >= '0' && tokens[i] <= '9')
            {
                StringBuffer sbuf = new StringBuffer();
                while (i < tokens.length && ((tokens[i] >= '0' && tokens[i] <= '9') || tokens[i] == '.'))
                    sbuf.append(tokens[i++]);
                values.push(Double.parseDouble(sbuf.toString().trim()));
                i--;
            }
            else if (Character.isLetter(tokens[i]) && tokens[i] != '+' && tokens[i] != '-' &&
                    tokens[i] != '*' && tokens[i] != '/' && tokens[i] != '(' && tokens[i] != ')')
            {
                StringBuffer sbuf = new StringBuffer();
                while (i < tokens.length && ((Character.isLetter(tokens[i]) && tokens[i] != '+' && tokens[i] != '-' && tokens[i] != '*' && tokens[i] != '/' && tokens[i] != '(' && tokens[i] != ')') ||
                        (tokens[i] == ' ' && (tokens[i+1] != '+' || tokens[i+1] != '-' || tokens[i+1] != '*' || tokens[i+1] != '/' || tokens[i+1] != '(' || tokens[i+1] != ')')) ||
                        (tokens[i] == ' ' && (tokens[i-1] != '+' || tokens[i-1] != '-' || tokens[i-1] != '*' || tokens[i-1] != '/' || tokens[i-1] != '(' || tokens[i-1] != ')')) ||
                        (Character.isDigit(tokens[i]) && (Character.isLetter(tokens[i-1]) || tokens[i] == ',' || Character.isDigit(tokens[i-1]) || tokens[i-1] == ' ')) || tokens[i] == '|' || tokens[i] == '_' || tokens[i] == '$' || tokens[i] == '%')) {
                    sbuf.append(tokens[i++]); }
                if(source.equals("Report Intersection")) {
                    Double reportData = getDataFromSummaryReport(reportTables, sbuf.toString(), reportDIMTable, val, isInputFormValidation,
                            isBackendReportValidation, valResults, bval);
                    if (reportData != null) {
                        values.push(reportData);
                    } else {
                        isReportMappingAvailable = false;
                    }
                } else if(source.equals("Input Form Intersection")) {
                    String eachIntersecName = sbuf.toString();
                    Double inputFormData = 0.0;
                    if (eachIntersecName.startsWith(reportSheetName.trim())) {
                        inputFormData = getDataFromInputFile(reportTables, eachIntersecName, reportDIMTable);
                        if (inputFormData != null) {
                            values.push(inputFormData);
                        } else {
                            isInputFormMappingAvailable = false;
                        }
                    } else {
                        for (DependencyWbData eachItem : dependencyDataList) {
                            File inputFile = eachItem.getDependencyWb();
                            List<String> worksheets = eachItem.getDependencyWsNamesList();
                            for (String worksheet : worksheets) {
                                if (eachIntersecName.startsWith(worksheet)) {
                                    String fileName = inputFile.getName().contains("(") ? inputFile.getName().split("\\(")[0].trim() : inputFile.getName().split("\\.")[0].trim();
                                    Table inputFileDIMTable = inputFormDIMTable.get(fileName + worksheet);
                                    Map<String, List<Object>> inputFileTables = inputFormTables.get(fileName + worksheet);
                                    inputFormData = getDataFromInputFile(inputFileTables, eachIntersecName, inputFileDIMTable);
                                    if (inputFormData != null) {
                                        values.push(inputFormData);
                                    } else {
                                        isInputFormMappingAvailable = false;
                                    }
                                }
                            }
                        }
                    }
                }
                i--;
            } else if (tokens[i] == '(') {
                ops.push(tokens[i]);
            } else if (tokens[i] == ')') {
                while (ops.peek() != '(')
                    values.push(applyOp(ops.pop(), values.pop(), values.pop()));
                ops.pop();
            } else if (tokens[i] == '+' || tokens[i] == '-' ||
                    tokens[i] == '*' || tokens[i] == '/' || tokens[i] == '.') {
                while (!ops.empty() && hasPrecedence(tokens[i], ops.peek()))
                    values.push(applyOp(ops.pop(), values.pop(), values.pop()));
                ops.push(tokens[i]);
            }
        }
        while (!ops.empty())
            values.push(applyOp(ops.pop(), values.pop(), values.pop()));
        resultList.add(values.pop());
        resultList.add(isReportMappingAvailable);
        resultList.add(isInputFormMappingAvailable);
        return resultList;
    }

    //Checks for math operations in the mapping document
    public static boolean hasPrecedence(char op1, char op2)
    {
        if (op2 == '(' || op2 == ')')
            return false;
        if ((op1 == '*' || op1 == '/') && (op2 == '+' || op2 == '-'))
            return false;
        else
            return true;
    }

    //Performs the calculations in the mapping document
    public static double applyOp(char op, double b, double a)
    {
        switch (op)
        {
            case '+':
                return a + b;
            case '-':
                return a - b;
            case '*':
                return a * b;
            case '/':
                if (b == 0)
                    throw new
                            UnsupportedOperationException("Cannot divide by zero");
                return a / b;
        }
        return 0.0;
    }

    //Reads the data from the Summary Report and returns the values
    public static Double getDataFromSummaryReport(Map<String, List<Object>> reportTables, String reportIntersectionName, Table reportDIMTable,
                                                  InputFormValidationData val, boolean isInputFormValidation, boolean isBackendReportValidation, ValidationDataResults valResults, BackendReportValidationData bval) {
        Table summaryselectTable = reportDIMTable.where(reportDIMTable.stringColumn("IntersectionName").isEqualTo(reportIntersectionName));
        Double summaryReportValue = Double.valueOf(0);
        if(summaryselectTable != null && summaryselectTable.rowCount() >0) {
            String tableName = summaryselectTable.row(0).getString("Table");
            valResults.setTableName(tableName);
            if(isInputFormValidation) {
                val.setTableName(tableName);
            } else if(isBackendReportValidation) {
                bval.setTableName(tableName);
            }
            List<Object> tableDetails = reportTables.get(tableName);
            Table summaryReportFilter = (Table) tableDetails.get(1);
            List<String> columnsWithPercentage = (List<String>) tableDetails.get(0);
            String selectColumn = summaryselectTable.row(0).getString("Select");
            if(isInputFormValidation) {
                val.setSelect(selectColumn);
            } else if(isBackendReportValidation) {
                bval.setSelect(selectColumn);
            }
            for (Column<?> column : summaryselectTable.columns()) {
                if (column.name().startsWith("Column") && !summaryselectTable.row(0).getString(column.name()).equals("")) {

                    String selectItem = summaryselectTable.row(0).getString(summaryselectTable.columnIndex(column.name()));
                    String selectValue = summaryselectTable.row(0).getString(summaryselectTable.columnIndex(column.name()) + 1);
                    if(column.name().equals("Column1")) {
                        if(isInputFormValidation) {
                            val.setColumn1Value(selectValue);
                            val.setColumn1Name(selectItem);
                        } else if(isBackendReportValidation) {
                            bval.setColumn1Value(selectValue);
                            bval.setColumn1Name(selectItem);
                        }
                    }
                    else if(column.name().equals("Column2")) {
                        if(isInputFormValidation) {
                            val.setColumn2Value(selectValue);
                            val.setColumn2Name(selectItem);
                        } else if(isBackendReportValidation) {
                            bval.setColumn2Value(selectValue);
                            bval.setColumn2Name(selectItem);
                        }
                    }
                    summaryReportFilter = summaryReportFilter.where(summaryReportFilter.stringColumn(selectItem)
                            .isEqualTo(selectValue));
                }
            }
            if(summaryReportFilter.select(selectColumn).get(0, 0) == null) {
                return null;
            }
            else if(summaryReportFilter.select(selectColumn).get(0, 0).getClass().getSimpleName().equalsIgnoreCase("Double")) {
                // summaryReportValue = (int) Math.round((Double) summaryReportFilter.select(selectColumn).get(0, 0));
                summaryReportValue = (Double) summaryReportFilter.select(selectColumn).get(0, 0);
                if(columnsWithPercentage.contains(selectColumn)) {
                    summaryReportValue = summaryReportValue * 100;
                }
            }
            else if(summaryReportFilter.select(selectColumn).get(0, 0).getClass().getSimpleName().equalsIgnoreCase("Integer")){
                summaryReportValue = Double.valueOf((int) summaryReportFilter.select(selectColumn).get(0, 0));
                if(columnsWithPercentage.contains(selectColumn)) {
                    summaryReportValue = summaryReportValue * 100;
                }
            }
        }
        else {
            return null;
        }
        return summaryReportValue;
    }

    //Reads the data from the input file and returns the values
    public static Double getDataFromInputFile(Map<String, List<Object>> inputFileTables, String inputIntersectionName, Table inputFileDIMTable) {
        Double data = Double.valueOf(0);
        Table inputSelectFilter = null;
        String selectInputColumn = null;
        Table inputSelectTable = inputFileDIMTable.where(inputFileDIMTable.stringColumn("IntersectionName").isEqualTo(inputIntersectionName));
        if (inputSelectTable != null && inputSelectTable.rowCount() > 0) {
            String tableName = inputSelectTable.row(0).getString("Table");
            List<Object> tableDetails = inputFileTables.get(tableName);
            inputSelectFilter = (Table) tableDetails.get(1);
            List<String> columnsWithPercentage = (List<String>) tableDetails.get(0);
            selectInputColumn = inputSelectTable.row(0).getString("Select");
            for (Column<?> column : inputSelectTable.columns()) {
                if (column.name().startsWith("Column") && !inputSelectTable.row(0).getString(column.name()).equals("")) {
                    String selectItem = inputSelectTable.row(0).getString(inputSelectTable.columnIndex(column.name()));
                    String selectValue = inputSelectTable.row(0).getString(inputSelectTable.columnIndex(column.name()) + 1);
                    inputSelectFilter = inputSelectFilter.where(inputSelectFilter.stringColumn(selectItem)
                            .isEqualTo(selectValue));
                }
            }
            if (inputSelectFilter != null) {
                if (inputSelectFilter.select(selectInputColumn).get(0, 0) == null) {
                    return null;
                } else if (inputSelectFilter != null) {
                    if (inputSelectFilter.select(selectInputColumn).get(0, 0).getClass().getSimpleName().equalsIgnoreCase("Double")) {
                        data = (Double) inputSelectFilter.select(selectInputColumn).get(0, 0);
                        if(columnsWithPercentage.contains(selectInputColumn)) {
                            data = data * 100;
                        }
                    } else if (inputSelectFilter.select(selectInputColumn).get(0, 0).getClass().getSimpleName().equalsIgnoreCase("Integer")) {
                        data = Double.valueOf((int) inputSelectFilter.select(selectInputColumn).get(0, 0));
                        if(columnsWithPercentage.contains(selectInputColumn)) {
                            data = data * 100;
                        }
                    }

                }
            } else {
                return null;
            }
        }
        return data;
    }
}
