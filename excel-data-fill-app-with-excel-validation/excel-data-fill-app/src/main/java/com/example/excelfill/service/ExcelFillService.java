package com.example.excelfill.service;

import com.example.excelfill.model.SheetMapping;
import com.example.excelfill.model.ValidationRule;
import com.fasterxml.jackson.databind.JsonNode;
import com.fasterxml.jackson.databind.ObjectMapper;
import org.apache.poi.ss.usermodel.*;
import org.apache.poi.ss.util.CellRangeAddress;
import org.apache.poi.ss.util.CellRangeAddressList;
import org.springframework.stereotype.Service;
import org.springframework.web.multipart.MultipartFile;

import java.io.ByteArrayOutputStream;
import java.io.InputStream;
import java.util.*;

@Service
public class ExcelFillService {

    private static final int HEADER_ROW_INDEX = 0;
    private static final int MAPPING_ROW_INDEX = 1;
    private static final int FIRST_DATA_ROW_INDEX = 2;

    private final ObjectMapper objectMapper = new ObjectMapper();
    private final DataFormatter dataFormatter = new DataFormatter();

    private static class ExcelValidationRule {
        private final DataValidationConstraint constraint;
        private final List<CellRangeAddress> ranges;
        ExcelValidationRule(DataValidationConstraint constraint, List<CellRangeAddress> ranges) {
            this.constraint = constraint;
            this.ranges = ranges;
        }
        DataValidationConstraint getConstraint() { return constraint; }
        List<CellRangeAddress> getRanges() { return ranges; }
    }

    public byte[] fillExcelTemplate(MultipartFile templateFile, MultipartFile jsonFile) throws Exception {
        JsonNode rootNode = objectMapper.readTree(jsonFile.getInputStream());
        try (InputStream inputStream = templateFile.getInputStream()) {
            return fillExcelTemplate(inputStream, rootNode);
        }
    }

    public byte[] fillExcelTemplate(InputStream templateInputStream, JsonNode rootNode) throws Exception {
        try (Workbook workbook = WorkbookFactory.create(templateInputStream);
             ByteArrayOutputStream outputStream = new ByteArrayOutputStream()) {

            for (int i = 0; i < workbook.getNumberOfSheets(); i++) {
                Sheet sheet = workbook.getSheetAt(i);
                Optional<SheetMapping> optionalMapping = readSheetMapping(sheet);
                if (optionalMapping.isEmpty()) continue;

                SheetMapping sheetMapping = optionalMapping.get();
                String objectName = sheetMapping.getObjectName();
                JsonNode dataArray = getChildIgnoreCase(rootNode, objectName);
                if (dataArray == null || !dataArray.isArray()) continue;

                Map<String, ValidationRule> validationRules = readValidationRules(rootNode, objectName);
                populateSheet(sheet, objectName, dataArray, sheetMapping.getColumnMappings(), validationRules);
            }

            workbook.write(outputStream);
            return outputStream.toByteArray();
        }
    }

    private Optional<SheetMapping> readSheetMapping(Sheet sheet) {
        Row mappingRow = sheet.getRow(MAPPING_ROW_INDEX);
        if (mappingRow == null) return Optional.empty();

        Map<Integer, String> mappings = new LinkedHashMap<>();
        String objectName = null;

        for (Cell cell : mappingRow) {
            String mapping = getCellStringValue(cell);
            if (mapping == null || mapping.isBlank() || !mapping.contains(".")) continue;

            mapping = mapping.trim();
            String currentObjectName = getObjectName(mapping);
            if (currentObjectName == null || currentObjectName.isBlank()) continue;

            if (objectName == null) objectName = currentObjectName;
            if (objectName.equalsIgnoreCase(currentObjectName)) mappings.put(cell.getColumnIndex(), mapping);
        }

        if (objectName == null || mappings.isEmpty()) return Optional.empty();
        return Optional.of(new SheetMapping(objectName, mappings));
    }

    private void populateSheet(Sheet sheet,
                               String objectName,
                               JsonNode dataArray,
                               Map<Integer, String> columnMappings,
                               Map<String, ValidationRule> validationRules) {

        int writeRowIndex = findNextAvailableRow(sheet);
        Row styleSourceRow = findStyleSourceRow(sheet);
        int firstInsertedRow = writeRowIndex;

        List<ExcelValidationRule> excelValidationRules = readExcelValidationRules(sheet);
        Map<Short, CellStyle> centeredStyleCache = new HashMap<>();

        for (JsonNode record : dataArray) {
            if (shouldSkipFullRow(record, objectName, columnMappings, validationRules)) continue;

            // This is the important part: validate against Excel template validations before adding row.
            if (!isRecordValidAgainstExcelValidations(record, objectName, columnMappings, excelValidationRules, writeRowIndex)) {
                continue;
            }

            Row currentRow = getOrCreateRow(sheet, writeRowIndex);
            copyTemplateRow(styleSourceRow, currentRow, columnMappings.keySet());

            for (Map.Entry<Integer, String> entry : columnMappings.entrySet()) {
                int columnIndex = entry.getKey();
                String mappingPath = entry.getValue();
                String fieldPath = removeObjectName(mappingPath);

                JsonNode valueNode = getValueByPath(record, fieldPath);
                Cell cell = getOrCreateCell(currentRow, columnIndex);
                copyCellStyle(styleSourceRow, cell, columnIndex, centeredStyleCache);

                ValidationRule rule = validationRules.get(toValidationKey(objectName, fieldPath));
                if (!isValueValid(valueNode, rule)) {
                    cell.setBlank();
                    continue;
                }

                setCellValue(cell, valueNode);
            }

            writeRowIndex++;
        }

        int lastInsertedRow = writeRowIndex - 1;
        if (lastInsertedRow >= firstInsertedRow) {
            applyInputValidationRules(sheet, objectName, columnMappings, validationRules, firstInsertedRow, lastInsertedRow);
        }
    }

    private List<ExcelValidationRule> readExcelValidationRules(Sheet sheet) {
        List<ExcelValidationRule> rules = new ArrayList<>();
        List<? extends DataValidation> validations = sheet.getDataValidations();
        if (validations == null || validations.isEmpty()) return rules;

        for (DataValidation validation : validations) {
            DataValidationConstraint constraint = validation.getValidationConstraint();
            if (constraint == null || validation.getRegions() == null) continue;

            CellRangeAddress[] addresses = validation.getRegions().getCellRangeAddresses();
            if (addresses == null || addresses.length == 0) continue;

            rules.add(new ExcelValidationRule(constraint, Arrays.asList(addresses)));
        }
        return rules;
    }

    private boolean isRecordValidAgainstExcelValidations(JsonNode record,
                                                         String objectName,
                                                         Map<Integer, String> columnMappings,
                                                         List<ExcelValidationRule> excelValidationRules,
                                                         int targetRowIndex) {
        if (excelValidationRules == null || excelValidationRules.isEmpty()) return true;

        for (Map.Entry<Integer, String> entry : columnMappings.entrySet()) {
            int columnIndex = entry.getKey();
            String fieldPath = removeObjectName(entry.getValue());
            JsonNode valueNode = getValueByPath(record, fieldPath);

            for (ExcelValidationRule rule : excelValidationRules) {
                if (!isValidationApplicable(rule, targetRowIndex, columnIndex)) continue;

                if (!isValueValidAgainstExcelConstraint(valueNode, rule.getConstraint())) {
                    System.out.println("Skipping row due to Excel validation failure. Object=" + objectName
                            + ", Field=" + fieldPath
                            + ", Value=" + valueNode
                            + ", Row=" + (targetRowIndex + 1)
                            + ", Column=" + (columnIndex + 1));
                    return false;
                }
            }
        }
        return true;
    }

    private boolean isValidationApplicable(ExcelValidationRule rule, int rowIndex, int columnIndex) {
        for (CellRangeAddress range : rule.getRanges()) {
            boolean rowMatches = rowIndex >= range.getFirstRow() && rowIndex <= range.getLastRow();
            boolean columnMatches = columnIndex >= range.getFirstColumn() && columnIndex <= range.getLastColumn();
            if (rowMatches && columnMatches) return true;
        }
        return false;
    }

    private boolean isValueValidAgainstExcelConstraint(JsonNode valueNode, DataValidationConstraint constraint) {
        if (valueNode == null || valueNode.isNull() || valueNode.isMissingNode()) return true;

        int validationType = constraint.getValidationType();

        if (validationType == DataValidationConstraint.ValidationType.INTEGER
                || validationType == DataValidationConstraint.ValidationType.DECIMAL) {

            if (!valueNode.isNumber()) return false;
            double value = valueNode.asDouble();
            Double formula1 = parseDouble(constraint.getFormula1());
            Double formula2 = parseDouble(constraint.getFormula2());
            return validateNumericValue(value, constraint.getOperator(), formula1, formula2);
        }

        if (validationType == DataValidationConstraint.ValidationType.LIST) {
            String[] allowedValues = constraint.getExplicitListValues();
            if (allowedValues == null || allowedValues.length == 0) return true;
            String value = valueNode.asText();
            for (String allowedValue : allowedValues) {
                if (allowedValue.equalsIgnoreCase(value)) return true;
            }
            return false;
        }

        if (validationType == DataValidationConstraint.ValidationType.TEXT_LENGTH) {
            String value = valueNode.asText();
            Double formula1 = parseDouble(constraint.getFormula1());
            Double formula2 = parseDouble(constraint.getFormula2());
            return validateNumericValue(value.length(), constraint.getOperator(), formula1, formula2);
        }

        return true;
    }

    private boolean validateNumericValue(double value, int operator, Double formula1, Double formula2) {
        if (formula1 == null && formula2 == null) return true;

        return switch (operator) {
            case DataValidationConstraint.OperatorType.BETWEEN -> formula1 != null && formula2 != null && value >= formula1 && value <= formula2;
            case DataValidationConstraint.OperatorType.NOT_BETWEEN -> formula1 != null && formula2 != null && (value < formula1 || value > formula2);
            case DataValidationConstraint.OperatorType.EQUAL -> formula1 != null && Double.compare(value, formula1) == 0;
            case DataValidationConstraint.OperatorType.NOT_EQUAL -> formula1 != null && Double.compare(value, formula1) != 0;
            case DataValidationConstraint.OperatorType.GREATER_THAN -> formula1 != null && value > formula1;
            case DataValidationConstraint.OperatorType.LESS_THAN -> formula1 != null && value < formula1;
            case DataValidationConstraint.OperatorType.GREATER_OR_EQUAL -> formula1 != null && value >= formula1;
            case DataValidationConstraint.OperatorType.LESS_OR_EQUAL -> formula1 != null && value <= formula1;
            default -> true;
        };
    }

    private Double parseDouble(String value) {
        if (value == null || value.isBlank()) return null;
        try {
            return Double.parseDouble(value.replace("\"", "").trim());
        } catch (Exception e) {
            return null;
        }
    }

    private void applyInputValidationRules(Sheet sheet,
                                           String objectName,
                                           Map<Integer, String> columnMappings,
                                           Map<String, ValidationRule> validationRules,
                                           int firstRow,
                                           int lastRow) {
        DataValidationHelper helper = sheet.getDataValidationHelper();

        for (Map.Entry<Integer, String> entry : columnMappings.entrySet()) {
            int columnIndex = entry.getKey();
            String fieldPath = removeObjectName(entry.getValue());
            ValidationRule rule = validationRules.get(toValidationKey(objectName, fieldPath));
            if (rule == null || rule.getType() == null) continue;

            DataValidationConstraint constraint = createConstraint(helper, rule);
            if (constraint == null) continue;

            CellRangeAddressList range = new CellRangeAddressList(firstRow, lastRow, columnIndex, columnIndex);
            DataValidation validation = helper.createValidation(constraint, range);
            validation.setSuppressDropDownArrow(false);
            validation.setShowErrorBox(true);
            if (rule.getErrorTitle() != null || rule.getErrorMessage() != null) {
                validation.createErrorBox(defaultText(rule.getErrorTitle(), "Invalid value"),
                        defaultText(rule.getErrorMessage(), "Value does not match the configured rule"));
            }
            sheet.addValidationData(validation);
        }
    }

    private DataValidationConstraint createConstraint(DataValidationHelper helper, ValidationRule rule) {
        String type = rule.getType();
        if (type == null) return null;
        if ("LIST".equalsIgnoreCase(type) && rule.getAllowedValues() != null && !rule.getAllowedValues().isEmpty()) {
            return helper.createExplicitListConstraint(rule.getAllowedValues().toArray(new String[0]));
        }
        String min = rule.getMin() == null ? null : String.valueOf(rule.getMin());
        String max = rule.getMax() == null ? null : String.valueOf(rule.getMax());
        if ("INTEGER".equalsIgnoreCase(type)) return helper.createIntegerConstraint(getOperator(rule), min, max);
        if ("NUMBER".equalsIgnoreCase(type) || "DECIMAL".equalsIgnoreCase(type)) return helper.createDecimalConstraint(getOperator(rule), min, max);
        if ("TEXT_LENGTH".equalsIgnoreCase(type)) return helper.createTextLengthConstraint(getOperator(rule), min, max);
        return null;
    }

    private int getOperator(ValidationRule rule) {
        if (rule.getMin() != null && rule.getMax() != null) return DataValidationConstraint.OperatorType.BETWEEN;
        if (rule.getMin() != null) return DataValidationConstraint.OperatorType.GREATER_OR_EQUAL;
        if (rule.getMax() != null) return DataValidationConstraint.OperatorType.LESS_OR_EQUAL;
        return DataValidationConstraint.OperatorType.IGNORED;
    }

    private Map<String, ValidationRule> readValidationRules(JsonNode rootNode, String objectName) {
        Map<String, ValidationRule> validationRules = new HashMap<>();
        JsonNode validationsNode = rootNode.path("validations");
        JsonNode node = getChildIgnoreCase(validationsNode, objectName);
        if (node == null || !node.isObject()) return validationRules;

        node.fields().forEachRemaining(entry -> {
            String fieldPath = entry.getKey();
            ValidationRule rule = objectMapper.convertValue(entry.getValue(), ValidationRule.class);
            validationRules.put(toValidationKey(objectName, fieldPath), rule);
        });
        return validationRules;
    }

    private boolean shouldSkipFullRow(JsonNode record,
                                      String objectName,
                                      Map<Integer, String> columnMappings,
                                      Map<String, ValidationRule> validationRules) {
        for (String mappingPath : columnMappings.values()) {
            String fieldPath = removeObjectName(mappingPath);
            ValidationRule rule = validationRules.get(toValidationKey(objectName, fieldPath));
            if (rule == null || !"ROW".equalsIgnoreCase(rule.getSkipMode())) continue;
            JsonNode valueNode = getValueByPath(record, fieldPath);
            if (!isValueValid(valueNode, rule)) return true;
        }
        return false;
    }

    private boolean isValueValid(JsonNode valueNode, ValidationRule rule) {
        if (rule == null) return true;
        if (valueNode == null || valueNode.isMissingNode() || valueNode.isNull()) return true;
        String type = rule.getType();
        if (type == null || type.isBlank()) return true;

        if ("INTEGER".equalsIgnoreCase(type)) {
            if (!valueNode.canConvertToLong()) return false;
            return isNumberInRange(valueNode.asDouble(), rule);
        }
        if ("NUMBER".equalsIgnoreCase(type) || "DECIMAL".equalsIgnoreCase(type)) {
            if (!valueNode.isNumber()) return false;
            return isNumberInRange(valueNode.asDouble(), rule);
        }
        if ("STRING".equalsIgnoreCase(type)) return valueNode.isTextual();
        if ("BOOLEAN".equalsIgnoreCase(type)) return valueNode.isBoolean();
        if ("LIST".equalsIgnoreCase(type)) return rule.getAllowedValues() == null || rule.getAllowedValues().contains(valueNode.asText());
        return true;
    }

    private boolean isNumberInRange(double value, ValidationRule rule) {
        if (rule.getMin() != null && value < rule.getMin()) return false;
        if (rule.getMax() != null && value > rule.getMax()) return false;
        return true;
    }

    private JsonNode getValueByPath(JsonNode record, String fieldPath) {
        if (record == null || fieldPath == null || fieldPath.isBlank()) return null;
        JsonNode current = record;
        for (String part : fieldPath.split("\\.")) {
            if (current == null || current.isNull() || current.isMissingNode()) return null;
            current = getChildIgnoreCase(current, part);
        }
        return current;
    }

    private JsonNode getChildIgnoreCase(JsonNode node, String fieldName) {
        if (node == null || fieldName == null || !node.isObject()) return null;
        JsonNode direct = node.get(fieldName);
        if (direct != null) return direct;
        Iterator<String> names = node.fieldNames();
        while (names.hasNext()) {
            String name = names.next();
            if (name.equalsIgnoreCase(fieldName)) return node.get(name);
        }
        return null;
    }

    private int findNextAvailableRow(Sheet sheet) {
        for (int rowIndex = FIRST_DATA_ROW_INDEX; rowIndex <= sheet.getLastRowNum(); rowIndex++) {
            Row row = sheet.getRow(rowIndex);
            if (row == null || isRowEmpty(row)) return rowIndex;
        }
        return Math.max(sheet.getLastRowNum() + 1, FIRST_DATA_ROW_INDEX);
    }

    private boolean isRowEmpty(Row row) {
        if (row == null) return true;
        if (row.getFirstCellNum() < 0 || row.getLastCellNum() < 0) return true;
        for (int i = row.getFirstCellNum(); i < row.getLastCellNum(); i++) {
            Cell cell = row.getCell(i);
            if (cell == null) continue;
            if (cell.getCellType() == CellType.FORMULA) return false;
            String value = getCellStringValue(cell);
            if (value != null && !value.isBlank()) return false;
        }
        return true;
    }

    private Row findStyleSourceRow(Sheet sheet) {
        Row firstDataRow = sheet.getRow(FIRST_DATA_ROW_INDEX);
        if (firstDataRow != null) return firstDataRow;
        Row mappingRow = sheet.getRow(MAPPING_ROW_INDEX);
        if (mappingRow != null) return mappingRow;
        return sheet.getRow(HEADER_ROW_INDEX);
    }

    private void copyTemplateRow(Row sourceRow, Row targetRow, Set<Integer> mappedColumns) {
        if (sourceRow == null || targetRow == null) return;
        targetRow.setHeight(sourceRow.getHeight());
        for (Cell sourceCell : sourceRow) {
            Cell targetCell = getOrCreateCell(targetRow, sourceCell.getColumnIndex());
            targetCell.setCellStyle(sourceCell.getCellStyle());
            if (!mappedColumns.contains(sourceCell.getColumnIndex())) copyCellValueOrFormula(sourceCell, targetCell);
        }
    }

    private void copyCellStyle(Row sourceRow, Cell targetCell, int columnIndex, Map<Short, CellStyle> centeredStyleCache) {
        if (sourceRow == null || targetCell == null) return;
        Cell sourceCell = sourceRow.getCell(columnIndex);
        if (sourceCell == null) return;

        CellStyle sourceStyle = sourceCell.getCellStyle();
        short styleIndex = sourceStyle.getIndex();
        CellStyle centeredStyle = centeredStyleCache.get(styleIndex);
        if (centeredStyle == null) {
            Workbook workbook = targetCell.getSheet().getWorkbook();
            centeredStyle = workbook.createCellStyle();
            centeredStyle.cloneStyleFrom(sourceStyle);
            centeredStyle.setAlignment(HorizontalAlignment.CENTER);
            centeredStyle.setVerticalAlignment(VerticalAlignment.CENTER);
            centeredStyleCache.put(styleIndex, centeredStyle);
        }
        targetCell.setCellStyle(centeredStyle);
    }

    private void copyCellValueOrFormula(Cell sourceCell, Cell targetCell) {
        if (sourceCell == null || targetCell == null) return;
        switch (sourceCell.getCellType()) {
            case FORMULA -> targetCell.setCellFormula(sourceCell.getCellFormula());
            case STRING -> targetCell.setCellValue(sourceCell.getStringCellValue());
            case NUMERIC -> targetCell.setCellValue(sourceCell.getNumericCellValue());
            case BOOLEAN -> targetCell.setCellValue(sourceCell.getBooleanCellValue());
            case BLANK -> targetCell.setBlank();
            default -> targetCell.setBlank();
        }
    }

    private void setCellValue(Cell cell, JsonNode valueNode) {
        if (valueNode == null || valueNode.isNull() || valueNode.isMissingNode()) {
            cell.setBlank();
            return;
        }
        if (valueNode.isNumber()) cell.setCellValue(valueNode.asDouble());
        else if (valueNode.isBoolean()) cell.setCellValue(valueNode.asBoolean());
        else cell.setCellValue(valueNode.asText());
    }

    private Row getOrCreateRow(Sheet sheet, int rowIndex) {
        Row row = sheet.getRow(rowIndex);
        return row != null ? row : sheet.createRow(rowIndex);
    }

    private Cell getOrCreateCell(Row row, int columnIndex) {
        Cell cell = row.getCell(columnIndex);
        return cell != null ? cell : row.createCell(columnIndex);
    }

    private String getCellStringValue(Cell cell) {
        if (cell == null) return null;
        return dataFormatter.formatCellValue(cell);
    }

    private String getObjectName(String mappingPath) {
        if (mappingPath == null || !mappingPath.contains(".")) return null;
        return mappingPath.split("\\.", 2)[0].trim();
    }

    private String removeObjectName(String mappingPath) {
        if (mappingPath == null || !mappingPath.contains(".")) return mappingPath;
        return mappingPath.split("\\.", 2)[1].trim();
    }

    private String toValidationKey(String objectName, String fieldPath) {
        return objectName + "." + fieldPath;
    }

    private String defaultText(String value, String defaultValue) {
        return value == null || value.isBlank() ? defaultValue : value;
    }
}
