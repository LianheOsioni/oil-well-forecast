# Oil Well Test-Separator Forecast Engine

A comprehensive VBA-based system for generating plausible historical-pattern simulations of 15-minute oil-well test-separator data (oil rates, gas rates, and BS&W values).

## Overview

This system learns statistical characteristics from historical 15-minute production intervals and generates future scenarios that preserve:
- **Mean and volatility**: Average production rate and natural variation
- **Autocorrelation**: Dependence on previous period's value
- **Regimes**: Low, normal, and high production periods
- **Events**: Sudden plunges and spikes
- **Targets**: Optional final values to reach (single or dual targets)
- **Multi-variable**: Simultaneous simulation of oil, gas, and BS&W

## Features

### Core Capabilities

1. **Historical Analysis**
   - Mean, median, standard deviation, quartiles
   - Trend direction and strength
   - Autocorrelation (lag-1, lag-2, lag-3)
   - Event detection (plunges, spikes)
   - Volatility assessment

2. **Forecasting Methods**
   - Three-regime production model (low, normal, high)
   - AR(1) autocorrelation blending
   - Mean reversion dynamics
   - Plunge/spike injection
   - Target-based trajectory constraint

3. **Multi-Variable Support**
   - Independent oil rate generation
   - Independent gas rate generation
   - BS&W value generation with centrifuge-compliant rounding
   - Cumulative-value support (auto-calculate from rates)

4. **Target Modes**
   - **Mode 0**: Free simulation (no target)
   - **Mode 1**: Reach single final target
   - **Mode 2**: Bridge two intermediate targets

5. **Safety & Reliability**
   - Input validation on all user entries
   - Range overlap detection
   - Output range verification
   - Atomic write operations
   - Error logging and user-friendly messages
   - Excel state restoration after completion

6. **Visualization**
   - Automatic chart creation
   - Multi-series chart support
   - Chart positioning and management

7. **Configuration Persistence**
   - Save/load settings to hidden cells (Z1:Z20)
   - Quick re-run with previous parameters
   - Configurable decimal places
   - Optional random seed for reproducibility

## Module Structure

### ModConfiguration
Manages persistent settings, configuration loading/saving, and parameter defaults.

**Key Types**: `ConfigSettings`
**Key Functions**:
- `LoadConfiguration()` - Restore saved settings
- `SaveConfiguration()` - Persist current configuration
- `GetConfigSummary()` - Display configuration summary
- `ColumnLetterToIndex()` / `ColumnIndexToLetter()` - Column reference conversion

### ModValidation
Input validation, range checking, and error reporting.

**Key Types**: `ValidationResult`
**Key Functions**:
- `ValidateColumnLetter()` - Check column letter format
- `ValidateLearningRange()` - Validate learning data range
- `ValidateNumericRange()` - Check range contains only numeric data
- `CheckOutputOverwrite()` - Detect data conflicts
- `ValidateTarget()` - Check target feasibility

### ModUserInput
User interface dialogs with proper cancellation handling.

**Key Functions**:
- `GetColumnInput()` - Request column with validation
- `GetRowInput()` - Request row with validation
- `GetNumericInput()` - Request number with optional range check
- `GetYesNoResponse()` - Request yes/no confirmation
- `RequestMultiVariableConfiguration()` - Guide through variable selection
- `RequestTargetConfiguration()` - Guide through target setup

### ModDataExtraction
Load data from worksheet, handle rates vs. cumulative, validate data quality.

**Key Types**: `DataExtractionResult`
**Key Functions**:
- `ExtractDataRange()` - Load range into array
- `ConvertCumulativeToRates()` - Calculate increments
- `ConvertRatesToCumulative()` - Reverse calculation
- `ValidateDataArray()` - Check data quality

### ModStatistics
Comprehensive statistical analysis of historical behavior.

**Key Types**: `StatisticsResult`, `TrendResult`, `AutocorrelationResult`
**Key Functions**:
- `CalculateStatistics()` - Mean, stdev, quartiles, skewness, kurtosis
- `AnalyzeTrend()` - Linear regression slope/intercept
- `AnalyzeAutocorrelation()` - Lag-1/2/3 and persistence
- `DetectPlunges()` - Count downward deviation events
- `DetectSpikes()` - Count upward deviation events
- `CalculateVolatility()` - Volatility from period-to-period changes

### ModForecast
Core forecasting engine and increment generation.

**Key Types**: `ForecastSession`
**Key Functions**:
- `InitializeSession()` - Set up forecast state from learned statistics
- `GenerateIncrement()` - Create single 15-minute increment
- `ValidateIncrement()` - Check plausibility
- `GenerateForecastBatch()` - Generate multiple increments
- `SetRandomSeed()` / `ClearRandomSeed()` - Control randomness

### ModTargeting
Target-based trajectory generation and constraint handling.

**Key Types**: `TargetTrajectory`
**Key Functions**:
- `SetupTargetTrajectory()` - Configure single or dual target
- `ApplyTargetConstraint()` - Adjust increment toward target
- `CalculateIdealTrajectory()` - Linear interpolation between targets
- `CheckTargetFeasibility()` - Quick feasibility check
- `EstimateTargetReachability()` - Human-readable assessment

### ModBSW
BS&W (Basic Sediment & Water) handling with centrifuge-compliant rounding.

**Rounding Tiers**:
- 0 to 0.1: 0.025 increments
- 0.1 to 1.0: 0.05 increments
- 1.0 to 10.0: 1.0 increments

**Key Functions**:
- `GenerateBSWValue()` - Create plausible BS&W value
- `RoundBSWToStandard()` - Apply centrifuge rounding
- `ValidateBSWValue()` - Check compliance
- `AnalyzeBSWData()` - Statistical summary
- `AnalyzeBSWTrend()` - Detect upward/downward trend

### ModOutput
Safe output writing with validation and atomic transactions.

**Key Functions**:
- `ValidateOutputRange()` - Pre-write validation
- `CheckColumnForFormulas()` - Detect protected cells
- `WriteOutputData()` - Atomic write operation
- `WriteMultiVariableOutput()` - Multi-column write
- `VerifyOutputWritten()` - Post-write validation
- `GenerateOutputSummary()` - Summary message

### ModChart
Chart creation, management, and visualization.

**Key Functions**:
- `CreateOrUpdateChart()` - Create or replace chart
- `CreateMultiSeriesChart()` - Multi-variable charting
- `DeleteChart()` - Remove chart
- `PositionChart()` - Reposition chart

### ModMain
Main orchestration and workflow control.

**Entry Point**: `GenerateOilWellForecast()`

**Workflow Phases**:
1. User input and configuration
2. Data extraction and validation
3. Statistical analysis
4. Target configuration (optional)
5. Forecast generation
6. Output and charting

### ModErrorHandling
Global error handling and logging.

**Key Functions**:
- `EnableErrorLogging()` / `DisableErrorLogging()`
- `LogError()` - Write error to log sheet
- `ReportValidationError()` - Display error to user
- `GetUserFriendlyErrorMessage()` - Convert error codes to messages
- `PromptToRetry()` - Ask for retry after failure

## Usage

### Basic Workflow

1. **Prepare data**: Arrange historical 15-minute increment data in a column.
2. **Run macro**: Call `GenerateOilWellForecast()` from Excel.
3. **Follow prompts**:
   - Select learning range (historical data)
   - Select output column and start row
   - Specify number of rows to generate
   - Set decimal places
   - (Optional) Specify target value(s)
4. **Review results**: Forecast appears in output column; chart (if selected) displays on worksheet.

### Advanced Features

#### Multi-Variable Simulation

Select multiple variables during configuration:
- Oil rate (required if "Oil" selected)
- Gas rate (required if "Gas" selected)
- BS&W (required if "BS&W" selected)

Each variable is generated independently with its own learned behavior.

#### Target-Based Generation

**Mode 1 (Single Target)**:
- Specify final cumulative value to reach
- System generates increments that converge to target
- Intermediate values remain realistic

**Mode 2 (Dual Target)**:
- Specify intermediate value at a middle row
- Specify final value at end row
- System generates trajectory bridging both targets

#### Reproducible Forecasts

To generate the same forecast twice:
1. Note the random seed used (or set one yourself)
2. Load configuration on second run
3. Specify the same seed in configuration
4. Results will be identical

## Configuration Persistence

Settings are saved to hidden cells Z1:Z20:

| Cell | Parameter | Default |
|------|-----------|----------|
| Z1 | Lookup Column | 2 (B) |
| Z2 | Learn Start | 2 |
| Z3 | Learn End | 51 |
| Z4 | Output Column | 3 (C) |
| Z5 | Output Start | 52 |
| Z6 | Num New Rows | 96 |
| Z7 | Decimal Places | 2 |
| Z8 | Plunge Frequency | 0.03 |
| Z9-Z13 | Multi-variable columns | 0 (unused) |
| Z14-Z16 | Target settings | 0 (no target) |
| Z17-Z19 | Advanced options | False/0 |

## Mathematical Model

### Increment Generation

Each increment is generated as:

1. **Regime Selection** (probability-based):
   - Low: 25% → `Mean - 0.5·σ ± 0.35·σ`
   - Mid: 50% → `Mean ± 0.35·σ`
   - High: 25% → `Mean + 0.5·σ ± 0.35·σ`

2. **Autocorrelation Blending**:
   ```
   Increment = Mean + ρ·(Previous - Mean) + Noise
   ```
   where ρ ∈ [0.05, 0.95] is lag-1 autocorrelation.

3. **Mean Reversion**:
   ```
   Adjustment = -(Previous - Mean) × 0.05
   ```

4. **Event Injection** (conditional):
   - Plunge (if triggered): Subtract random depth
   - Spike (if triggered): Add random height

5. **Target Constraint** (if active):
   ```
   Blend = 0.15 + 0.5·(1 - RemainingRows/TotalRows)
   FinalIncrement = Blend·Required + (1-Blend)·Generated
   ```

### Statistical Estimates

- **Mean**: Simple average
- **StDev**: Sample standard deviation (n-1)
- **Autocorr**: Lag-1 ratio of covariance to variance
- **Bands**: Mean ± 0.5·σ
- **Plunge Threshold**: Mean - 1.5·σ
- **Spike Threshold**: Mean + 1.5·σ

## Error Codes

| Code | Meaning |
|------|----------|
| 1001 | Learning start row invalid |
| 1002 | Learning end row invalid |
| 1003 | Insufficient learning observations |
| 2001 | Range contains empty cells |
| 2002 | Range contains error cells |
| 2003 | Range contains non-numeric values |
| 3001 | Output overlaps learning data |
| 3002 | Output range has existing data |
| 4001 | Target not numeric |
| 4002/4003 | Target unrealistic |
| 5001 | Invalid output range |
| 5002 | Output would overwrite history |

## Testing

See `ModTesting.bas` for:
- Unit tests for each module
- Integration tests for complete workflow
- Example data sets
- Validation test cases

## Limitations & Future Enhancements

### Current Limitations
1. Single column learning (future: multi-column trends)
2. Fixed three-regime model (future: adaptive regime detection)
3. No seasonal pattern detection (future: periodic analysis)
4. No external variable support (future: multivariate forecasting)

### Planned Enhancements
1. Bootstrap confidence intervals
2. Scenario comparison and branching
3. Batch processing of multiple wells
4. Export to CSV/JSON
5. Real-time data feed integration
6. Web-based visualization

## Notes

- All values are rounded **after** generation to preserve mathematical integrity
- Cumulative values are calculated from increments, not rounded independently
- BS&W values use centrifuge-compliant rounding (0.025, 0.05, or 1.0 increments)
- The system assumes 15-minute intervals (96 per 24-hour day)
- Negative increments are allowed (for declining production)

---

**Author**: Copilot
**Version**: 1.0.0
**Last Updated**: 2026-09-06
