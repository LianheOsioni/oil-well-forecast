# Technical Assumptions & Methodology

## Data Assumptions

### Input Data Format

1. **Temporal Interval**: 15 minutes (96 intervals per 24-hour day)
2. **Data Type**: Numeric increments or cumulative values
3. **Range**: Contiguous rows, no missing values
4. **Precision**: Any decimal precision (will be rounded to specification)
5. **Sign**: Can be positive, negative, or zero

### Data Quality

- System assumes:
  - Data is measured/recorded accurately
  - Outliers are genuine events (not data entry errors)
  - No systematic sensor drift or calibration issues
  - Time stamps are regular and synchronized

- System does NOT:
  - Automatically detect/remove artificial spikes
  - Correct for systematic bias
  - Handle seasonality (daily/weekly patterns)
  - Account for external events or shutdowns

## Statistical Model

### Distribution Assumptions

**Historical Analysis**:
- Mean and standard deviation calculated as sample statistics (n-1)
- No assumption of normality (works with any distribution)
- Uses sample quantiles (not parametric quartiles)

**Forecast Generation**:
- Three regimes (low/normal/high) with fixed probabilities
- Autocorrelation model: AR(1) with blending
- Event injection: Plunges and spikes follow detected historical frequency
- No assumption of specific parametric distribution

### Mathematical Model

**Increment Generation Formula**:

```
Generated = f(Regime, Autocorr, MeanReversion, Events, Target)
```

Where:

1. **Regime Selection** (probability-based):
   - P(Low) = 25% → mean - 0.5σ ± 0.35σ
   - P(Normal) = 50% → mean ± 0.35σ
   - P(High) = 25% → mean + 0.5σ ± 0.35σ

2. **Autocorrelation Blending**:
   ```
   AR(1) = μ + ρ·(X_{t-1} - μ) + ε_t
   ```
   where ρ = lag-1 autocorrelation, ε_t ~ N(0, σ²)

3. **Mean Reversion Adjustment**:
   ```
   Drift = -(X_{t-1} - μ) × 0.05
   ```
   Pulls values toward historical mean with 5% strength

4. **Event Injection** (conditional):
   - Plunge: subtract random depth ~ U(0.4, 0.6)·D_avg
   - Spike: add random height ~ U(0.4, 0.6)·S_avg
   - Probability: user-configurable (default 0.03 each)

5. **Target Constraint** (if active):
   ```
   Blend = 0.15 + 0.5·(1 - r_remaining/r_total)
   Final = Blend·Required + (1-Blend)·Generated
   ```
   Gradually increases blending as target row approaches

### Autocorrelation Bounds

- Lag-1 is clipped to [0.05, 0.95]
  - Minimum ensures some randomness
  - Maximum prevents over-persistence
- Persistence score = 0.6·|ρ₁| + 0.3·|ρ₂| + 0.1·|ρ₃|
  - Ranges 0 (independent) to 1 (highly persistent)

### Plunge & Spike Parameters

- **Threshold**: Mean ± 1.5σ
  - Events outside this band are considered extreme
- **Frequency**: Calculated from historical count
  - Count(Events) / Total(Observations)
- **Average Depth**: Mean absolute deviation from threshold
  - For plunges: average by how far below threshold
  - For spikes: average by how far above threshold

## Validation & Safety Mechanisms

### Input Validation

1. **Column References**: A-XFD (16,384 max)
2. **Row References**: 1-1,048,576 (Excel limit)
3. **Numeric Ranges**: Any double precision (-10^308 to 10^308)
4. **Learning Data**: Minimum 10 observations required
5. **Output Configuration**: No overlap with learning range

### Data Integrity Checks

1. **Before Analysis**:
   - Range contains only numeric values (no text/errors)
   - No empty cells within learning range
   - Statistical parameters are valid (mean, σ, ρ)

2. **During Generation**:
   - Generated increment passes bounds checking
   - Cumulative value stays within 3σ of historical range
   - Target constraint doesn't force unrealistic increments

3. **Before Output**:
   - Output column doesn't overlap learning data
   - Output range has been validated
   - All array sizes match expected counts

4. **After Writing**:
   - Verification pass confirms all cells written
   - No data loss or corruption
   - Excel state restored (screen updates, formulas, events)

### Bounds Enforcement

**Generated Increment Limits**:
- Default: Mean ± 3σ
- User configurable via `MaxDeviationFactor`
- Hard floor: Must be positive if `AllowExtrapolation=False`

**Target Reachability**:
- Feasible: Required avg increment within Mean ± 3σ
- Challenging: Within Mean ± 5σ
- Unrealistic: Beyond Mean ± 5σ

## Rounding & Precision

### Standard Decimal Rounding

```vb
Rounded = Round(Value, DecimalPlaces)
```

- Banker's rounding (round half to even)
- Applied only for output display
- Mathematical calculations use full precision

### BS&W Special Rounding

Three-tier system based on centrifuge standard:

| Range | Increment | Examples |
|-------|-----------|----------|
| 0 - 0.1 | 0.025 | 0, 0.025, 0.05, 0.075 |
| 0.1 - 1.0 | 0.05 | 0.1, 0.15, 0.2, ..., 0.95 |
| 1.0 - 10.0 | 1.0 | 1, 2, 3, ..., 10 |
| > 10.0 | Capped | Max 10.0 |

**BS&W Rounding Algorithm**:
```
For each BS&W value:
  1. If > 10.0: cap at 10.0
  2. Determine tier (0.1, 1.0, or 10.0 increments)
  3. Round to nearest increment in that tier
  4. Apply tier-specific rounding function
```

## Target Mode Methodology

### Mode 0: Free Simulation

- Pure regime and AR(1) model
- No constraint toward any value
- Most realistic for exploratory forecasting
- No guarantee on final value

### Mode 1: Single Target

- Specify: `Target1` and `Target1Row`
- Required average increment: (Target1 - Current) / Target1Row
- Blending factor increases from 15% to 90% as target row approaches
- Ensures target is reached (within rounding error)

### Mode 2: Dual Target

- Specify: `Target2` at `Target2Row`, then `Target1` at `Target1Row`
- Two-segment trajectory:
  - Segment 1: Current → Target2 (rows 1 to Target2Row)
  - Segment 2: Target2 → Target1 (rows Target2Row+1 to Target1Row)
- Each segment applies independent blending

**Dual Target Feasibility**:
- Target2Row must be < Target1Row
- Both targets must be feasible independently
- System warns if path requires extreme increments

## Event Detection Algorithm

### Plunge Detection

```
Threshold = Mean - 1.5·σ
For each observation:
  If value < Threshold:
    Count as plunge
    Depth += |Mean - value|
```

**Frequency**: Count / Total Observations  
**Average Depth**: Sum(Depths) / Count(Plunges)

### Spike Detection

```
Threshold = Mean + 1.5·σ
For each observation:
  If value > Threshold:
    Count as spike
    Height += |value - Mean|
```

**Frequency**: Count / Total Observations  
**Average Height**: Sum(Heights) / Count(Spikes)

### Event Injection in Generation

At each step:
```
IF Rand() < PlungeFrequency THEN
  Plunge_Depth = Rand(0.4, 0.6) × Average_Plunge_Depth
  Increment -= Plunge_Depth
END IF

IF Rand() < SpikeFrequency THEN
  Spike_Height = Rand(0.4, 0.6) × Average_Spike_Height
  Increment += Spike_Height
END IF
```

**Note**: Plunge and spike can occur in same interval (independent draws)

## Cumulative vs. Rate Handling

### Rate Data (Default)

Input: Column of 15-minute increments  
Output: Cumulative by summing increments

```
Cumulative(t) = CumulativeLastRow + Sum(Forecast_1 to Forecast_t)
```

### Cumulative Data (Optional)

If user provides cumulative values:
1. System auto-detects (differences are rates)
2. Converts to rates: `Rate(i) = Cumulative(i) - Cumulative(i-1)`
3. Learns from rates
4. Forecasts rates
5. Converts back to cumulative for output

## Excel Interaction Assumptions

### Workbook State

- System assumes workbook is open and active
- User has permission to read/write cells
- No worksheet protection in effect
- Calculation mode can be temporarily changed

### Performance Optimization

During forecast run:
- Screen updating disabled (faster writes)
- Calculation set to Manual (avoid intermediate recalcs)
- Event handling disabled (no VB event fires)
- **Restored after completion** (automatic)

### Chart Support

- Requires Excel 2016+ for full chart features
- Charts are embedded in worksheet (not separate sheets)
- Chart type: XY Scatter with lines and markers
- Chart includes legend and gridlines

## Random Number Generation

### Randomness Source

```vb
Rnd()  ' Built-in VBA random number generator
```

- Returns uniform [0, 1) pseudo-random values
- Seeded by system clock (or `Randomize seed`)
- Suitable for Monte Carlo simulation
- **NOT cryptographically secure** (not relevant here)

### Reproducibility

```vb
Randomize 12345  ' Set seed before generation
' Now Rnd() produces same sequence as other runs with seed 12345
```

**Use Case**: Frozen scenario for reporting or comparison

### Approximate Normal Distribution

```vb
Noise = (Rnd + Rnd + Rnd + Rnd - 2) * σ  ' Approximate N(0, σ)
```

Sum of 4 uniforms ≈ normal (Central Limit Theorem)  
Mean of 4: 2.0, so subtract 2 to center at 0

## Limitations & Caveats

### Known Limitations

1. **Single-Column Learning**
   - Only learns from one historical column
   - Cannot capture multi-variable correlations
   - Future: Multi-column trend analysis

2. **Static Regime Probabilities**
   - Low/Normal/High always 25/50/25
   - Doesn't adapt if regime changes over time
   - Future: Adaptive regime detection

3. **No Seasonality**
   - Doesn't detect or project daily/weekly cycles
   - Mean is constant (ignores trends)
   - Future: Periodic component analysis

4. **Event Frequencies**
   - Fixed based on historical count
   - Assumes future events similar to past
   - Doesn't account for increasing/decreasing trend

5. **Target Constraint Blending**
   - Fixed blending schedule (linear increase)
   - May require extreme increments at end
   - Future: Nonlinear or adaptive blending

### When NOT to Use

- **High-precision safety-critical applications**: This is a heuristic tool, not engineering simulation
- **Regime shift analysis**: System assumes regime structure stable
- **Long-term forecasting**: Beyond 7-10 days, model confidence drops
- **External shock scenarios**: Cannot model sudden well changes or equipment failure
- **Regulatory compliance forecasting**: Verify with domain experts before regulatory use

## Best Practices

1. **Learning Range**: Use 500-1000 historical observations if available
   - Minimum 50 (system enforces 10)
   - Diminishing returns beyond 2000

2. **Validation**:
   - Compare multiple target scenarios
   - Check if forecast range matches historical range
   - Verify autocorrelation is reasonable (0.1-0.7)

3. **Reporting**:
   - Document learning range used
   - Note target modes and values
   - Include statistical summary from analysis phase
   - Caveat: "Plausible scenarios for analysis purposes"

4. **Iteration**:
   - Run multiple forecasts with different seeds
   - Compare scenarios and distributions
   - Adjust target values if results unrealistic

---

**Last Updated**: 2026-09-06  
**Document Version**: 1.0
