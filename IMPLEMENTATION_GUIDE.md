# Implementation Guide

## Installation Instructions

### Step 1: Prepare Excel Workbook

1. Open Microsoft Excel (2016 or later recommended)
2. Create a new workbook or open an existing one with historical test-separator data
3. Ensure your data is organized with:
   - Header row in row 1
   - Historical 15-minute increments in a single column (e.g., column B)
   - Data rows contiguous with no gaps

### Step 2: Import VBA Modules

1. Open the Visual Basic Editor (Alt+F11)
2. In the Project Explorer, right-click on your workbook
3. Select "Import File"
4. Import each `.bas` file in this order:
   1. `ModConfiguration.bas`
   2. `ModValidation.bas`
   3. `ModUserInput.bas`
   4. `ModDataExtraction.bas`
   5. `ModStatistics.bas`
   6. `ModForecast.bas`
   7. `ModTargeting.bas`
   8. `ModBSW.bas`
   9. `ModOutput.bas`
   10. `ModChart.bas`
   11. `ModMain.bas`
   12. `ModErrorHandling.bas`
   13. `ModTesting.bas`

### Step 3: Create Entry Point

1. Still in the VBA editor, insert a new module: Insert → Module
2. Name it `ModUI`
3. Add this code:

```vb
Sub GenerateForecast()
    GenerateOilWellForecast
End Sub
```

4. Save the workbook as Excel Macro-Enabled (.xlsm)

### Step 4: Verify Installation

1. Go to Tools → Macros → Macros
2. Select `ValidateInstallation`
3. Click "Run"
4. Confirm all modules report installed successfully

## Quick Start

### Running Your First Forecast

1. Open your workbook with historical data
2. Tools → Macros → Macros → Select `GenerateForecast` → Run
3. Follow the on-screen prompts:
   - **Step 1**: Select learning data range (your historical 15-minute increments)
   - **Step 2**: Select output location (where forecast will be written)
   - **Step 3**: System loads and analyzes your historical data
   - **Step 4**: Review statistical summary
   - **Step 5**: Configure optional targets (or skip for free simulation)
   - **Step 6**: Review forecast results and optional chart

### Understanding the Output

The forecast generates a column of cumulative values that:
- Preserve the mean and volatility of your historical data
- Maintain autocorrelation from period to period
- Include realistic plunges and spikes
- Reach optional target values if configured

## Configuration Details

### Default Settings

| Parameter | Default | Meaning |
|-----------|---------|----------|
| Lookup Column | B | Historical data source |
| Learn Start | 2 | First row of historical data |
| Learn End | 51 | Last row (50 observations) |
| Output Column | C | Forecast destination |
| Output Start | 52 | First forecast row |
| Num New Rows | 96 | 24 hours of 15-min intervals |
| Decimal Places | 2 | Rounding precision |
| Plunge Frequency | 0.03 | ~3% chance per interval |
| Target Mode | 0 | No target (free simulation) |
| Allow Extrapolation | True | Values can exceed historical range |

### Customizing Configuration

Settings are saved in hidden cells Z1:Z20. To modify:

1. Tools → Options → View (check "Hidden and empty cell markers")
2. Unhide rows/columns containing Z1:Z20
3. Edit values directly
4. Save workbook
5. Run forecast again (loads your custom settings)

## Advanced Usage

### Multi-Variable Forecasting

**For Oil + Gas + BS&W:**

1. Prepare three data columns with historical rates
2. Run forecast for each variable separately, OR
3. Modify `PhaseUserInput()` in ModMain to enable simultaneous multi-variable mode

**To enable multi-variable:**

- In ModMain, uncomment: `RequestMultiVariableConfiguration config`
- Follow prompts to specify columns for each variable
- System generates independent forecasts for each

### Target-Based Scenarios

**Example: Decline Analysis**

If you want to generate data that declines from current 150 to a target 50 over 96 intervals:

1. Run forecast normally
2. When prompted for targets, select "Yes"
3. Enter Target 1 = 50
4. Target 1 Row = 96
5. System generates path from 150 → 50

**Example: Multi-Phase Decline**

Declining from 150 → 120 (in first 48 rows) → 80 (in final 96 rows):

1. Select "Two Targets"
2. Target 2 = 120, Row = 48
3. Target 1 = 80, Row = 96
4. System generates two-segment trajectory

### Reproducible Forecasts

To ensure identical results on repeated runs:

1. Note the random seed from first run
2. Open configuration cells Z18-Z19
3. Set Z18 = 1 (use seed)
4. Set Z19 = [your seed number]
5. Run forecast again

## Troubleshooting

### Issue: "Error 1003: Insufficient learning observations"

**Cause**: Learning range has fewer than 10 rows  
**Solution**: Select a larger historical range (minimum 10-20 rows recommended)

### Issue: "Error 3001: Output overlaps learning data"

**Cause**: Output start row is within or before the learning range  
**Solution**: Set Output Start to a row after your learning data ends

### Issue: "Error 2001: Range contains empty cells"

**Cause**: Your selected column has missing values  
**Solution**: 
- Remove rows with empty cells, or
- Use a different column without gaps

### Issue: Chart doesn't appear

**Cause**: Chart creation failed due to Excel restrictions  
**Solution**: 
- Click "No" when asked about chart
- Manually create chart after forecast completes
- Verify Excel version supports charts in your configuration

### Issue: "Module not found" error

**Cause**: Not all modules were imported, or import order was wrong  
**Solution**: 
- Re-import all modules in the order listed
- Verify no duplicate module names
- Save and close workbook, reopen and retry

## Performance Considerations

### Data Load Time

- 50 observations: < 1 second
- 500 observations: 1-2 seconds
- 5000 observations: 5-10 seconds

*Time varies with hardware and Excel version*

### Memory Usage

- Array storage: ~8 bytes per numeric value
- 1000 forecast rows ≈ 16 KB
- Chart storage: ~50-200 KB

### Optimization Tips

1. **Batch Processing**: Generate multiple scenarios in separate worksheets
2. **Large Datasets**: Use first 500-1000 observations only (learning diminishes beyond this)
3. **Performance Mode**: Disable charting for fastest execution

## Safety & Backup

### Before Running Forecast

1. **Backup your workbook** (save as `.xlsx` first)
2. **Test on a copy** (verify output location won't overwrite important data)
3. **Review configuration** (output start row, number of rows, decimal places)

### Data Integrity

- Historical data is **never modified** by this system
- Output range is validated before any write
- User is prompted if data exists in output location
- All changes can be undone (Ctrl+Z) before saving

## Validation Checklist

Before each forecast run:

- [ ] Historical data is numeric (no text, no formulas)
- [ ] Learning range has no gaps or errors
- [ ] Output column is different from learning column (recommended)
- [ ] Output start row is after learning end row
- [ ] Decimal places is between 0-10
- [ ] Number of rows to generate is > 0
- [ ] (If targets): target values are realistic
- [ ] (If targets): target row is > 0 and ≤ NumNewRows

## Getting Help

### View Calculation Details

Modulation includes detailed comments explaining:
- Statistical formulas used
- Random number generation approach
- Regime transition logic
- Target constraint blending

Refer to:
- `ModStatistics.bas` for statistical methods
- `ModForecast.bas` for generation algorithm
- `ModTargeting.bas` for target-based constraints

### Test the System

1. Tools → Macros → Select `GenerateExampleData` → Run
2. This creates a sheet with 1440 synthetic observations
3. Run forecast on this data to verify installation

### Enable Error Logging

In ModMain, at the start of `GenerateOilWellForecast()`:

```vb
EnableErrorLogging  ' Add this line
```

Errors will be logged to hidden sheet `_ErrorLog`.

---

## Next Steps

1. ✓ Install modules (follow steps above)
2. ✓ Run test forecast on example data
3. ✓ Run forecast on your real historical data
4. ✓ Review results and adjust configuration as needed
5. ✓ Use output for further analysis, reporting, or decisions

---

**For detailed technical information, see README.md**
