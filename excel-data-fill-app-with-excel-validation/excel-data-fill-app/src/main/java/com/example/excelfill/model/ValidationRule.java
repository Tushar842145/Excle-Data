package com.example.excelfill.model;
import java.util.List;
public class ValidationRule {
    private String type;
    private Double min;
    private Double max;
    private String skipMode = "CELL";
    private String errorTitle;
    private String errorMessage;
    private List<String> allowedValues;
    public String getType() { return type; }
    public void setType(String type) { this.type = type; }
    public Double getMin() { return min; }
    public void setMin(Double min) { this.min = min; }
    public Double getMax() { return max; }
    public void setMax(Double max) { this.max = max; }
    public String getSkipMode() { return skipMode; }
    public void setSkipMode(String skipMode) { this.skipMode = skipMode; }
    public String getErrorTitle() { return errorTitle; }
    public void setErrorTitle(String errorTitle) { this.errorTitle = errorTitle; }
    public String getErrorMessage() { return errorMessage; }
    public void setErrorMessage(String errorMessage) { this.errorMessage = errorMessage; }
    public List<String> getAllowedValues() { return allowedValues; }
    public void setAllowedValues(List<String> allowedValues) { this.allowedValues = allowedValues; }
}
