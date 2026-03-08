# FBCJ Church Plant Analysis: Jackson vs Dalton

Geospatial analysis of FBCJ (First Baptist Church of Jackson) member addresses to evaluate a potential second campus in Dalton, OH. Answers the question: **"Which current members would be closer to a Dalton location, and how likely are they to prefer it based on driving time?"**

## Quick Start

```bash
pip install folium matplotlib numpy openpyxl plotly requests scipy
python3 church_map.py
```

On first run, the script geocodes addresses and fetches driving times from Google Maps (~30 min for API calls). Subsequent runs use cached data and complete in seconds.

## Outputs

| File | Description |
|------|-------------|
| `FBCJ_Church_Map.html` | Interactive map with member dots, equal-distance and equal-time boundary lines |
| `FBCJ_Church_Map_Time.html` | Contour map showing driving time advantage zones (red=Jackson closer, blue=Dalton closer) |
| `FBCJ_Church_Map_Distance.html` | Contour map showing driving distance advantage zones |
| `FBCJ_Church_Probability_Distance.html` | Scatter plot: distance to Jackson vs distance to Dalton per family |
| `FBCJ_Church_Probability_Time.html` | Scatter plot: driving time to Jackson vs driving time to Dalton per family |
| `FBCJ_Church_Histogram.html` | Distribution of driving time differences with KDE density curve |
| `FBCJ_Sensitivity_Analysis.html` | How conclusions change across different model parameters (K and incumbency bias) |
| `FBCJ_Analysis_Results.xlsx` | Full Excel export — Sheet 1: every family sorted by P(Dalton) desc; Sheet 2: summary + methodology |
| `geocoded_addresses.json` | Cached geocoding results and driving distances (avoids redundant API calls) |

## Data Pipeline

```
FBCJ_Directory #2.xlsx          (source: ServantKeeper church directory)
    → Geocode addresses          (US Census Bureau primary, Nominatim ZIP fallback)
    → Calculate driving times     (Google Maps Distance Matrix API, Sunday 9 AM departure)
    → Filter outliers             (>35 mi from church midpoint removed)
    → Compute probabilities       (logistic model on driving time difference)
    → Flag data quality           (mark approx/fallback entries, compute household sizes)
    → Statistical tests           (Wilcoxon, binomial test, effect size)
    → Sensitivity analysis        (K parameter + incumbency bias)
    → Generate maps + charts
    → Export Excel
```

## The Model

### Formula

```
P(Jackson) = 1 / (1 + exp(K × (time_to_Jackson − time_to_Dalton)))
```

This is a **logistic (sigmoid) function** that converts a driving time difference into a probability between 0 and 1.

### Inputs

- **time_to_Jackson**: Google Maps driving time from the family's address to Jackson (minutes)
- **time_to_Dalton**: Google Maps driving time from the family's address to Dalton (minutes)
- **time_to_Jackson − time_to_Dalton**: the time difference. Positive = Jackson takes longer. Negative = Dalton takes longer.

### The K Parameter

**K controls how sharply people respond to driving time differences.** Think of it as a "sensitivity dial."

- **Low K (0.05)**: People barely care about time. Even 15 minutes closer to Dalton only yields ~68% Dalton probability. Most families land in the toss-up zone.
- **High K (0.30)**: People are very time-sensitive. A 5-minute difference pushes probability to ~80%. Almost nobody is in the toss-up zone.
- **Current K (0.15)**: A moderate middle ground.

### Example Probabilities at K=0.15

| Scenario | Time Diff | P(Jackson) | Interpretation |
|----------|-----------|------------|----------------|
| Jackson is 10 min closer | −10 min | **82%** | Strongly lean Jackson |
| Jackson is 5 min closer | −5 min | **68%** | Moderately lean Jackson |
| Equal driving time | 0 min | **50%** | Coin flip |
| Dalton is 5 min closer | +5 min | **32%** | Moderately lean Dalton |
| Dalton is 10 min closer | +10 min | **18%** | Strongly lean Dalton |
| Dalton is 20 min closer | +20 min | **5%** | Almost certainly Dalton |

**Important:** K=0.15 is not calibrated from survey data — it is an assumed value. The sensitivity analysis shows how conclusions change across different K values. See "Sensitivity Analysis" below.

### Classification Thresholds

- **Lean Jackson**: P(Jackson) > 60%
- **Lean Dalton**: P(Jackson) < 40% (equivalently P(Dalton) > 60%)
- **Toss-up**: P(Jackson) between 40–60% (maps to roughly ±2.7 min time difference at K=0.15)

## Sensitivity Analysis

Because K is assumed (not measured), the sensitivity analysis tests whether conclusions hold across a range of reasonable values.

### K Parameter Sensitivity (left panel of chart)

- X-axis: K values from 0.05 to 0.30
- Y-axis: percentage of families in each category
- Three lines: Lean Jackson (red), Toss-up (gray), Lean Dalton (blue)

**How to read it:** If the red line is always above the blue line regardless of K, then "more families lean Jackson than Dalton" is a robust finding. The exact percentages depend on K, but the direction doesn't.

**K-independent findings** are reported separately — families who are >15 minutes closer to one church will lean that direction under any reasonable K. These are the most trustworthy results.

### Incumbency Bias Sensitivity (right panel of chart)

Current model assumes equidistant = 50/50 coin flip. In reality, people already attend Jackson — they have relationships, habits, kids in programs. They won't switch for a 1-minute advantage.

**Incumbency bias** shifts the break-even point: "How many extra minutes closer must Dalton be before it's a fair fight?"

- **Bias = 0** (current): Equal time = 50/50
- **Bias = 5**: Dalton must be 5 min closer to reach 50/50
- **Bias = 10**: Dalton must be 10 min closer to reach 50/50
- **Bias = 15**: Dalton must be 15 min closer to reach 50/50

The chart shows how the Jackson/Toss-up/Dalton split shifts as you increase bias. Church planting research suggests 15–20 minutes of advantage is typically needed to overcome inertia for established members.

## Statistical Tests

The analysis includes two statistical tests to confirm whether observed patterns are real or could be random chance.

### Wilcoxon Signed-Rank Test

**Question:** "Is the typical (median) time difference truly non-zero, or could these differences have come from a distribution centered at zero?"

- If p < 0.05: the median time difference is statistically significant — there's a real directional pattern.
- If p ≥ 0.05: the observed pattern could be random noise.

### Binomial Test

**Question:** "Is the Jackson-closer vs Dalton-closer split significantly different from 50/50?"

For example, if 62% of families are closer to Jackson: could you get that result by flipping a fair coin 300 times? The binomial test answers this.

### High-Quality vs All Data

Both tests are reported twice:
- **All addresses**: includes every geocoded family
- **High-quality only**: excludes families whose location or driving time was estimated (ZIP centroid or straight-line fallback)

This separation ensures that low-quality data points aren't driving the conclusions.

## Data Quality

### Geocoding Quality

| Level | Method | Accuracy |
|-------|--------|----------|
| **Exact** | US Census Bureau geocoder matched the street address | ~10 meters |
| **Approximate** | Nominatim ZIP code centroid (Census failed) | ~1–5 miles |

### Driving Time Quality

| Level | Method | Accuracy |
|-------|--------|----------|
| **Google Maps** | Distance Matrix API, Sunday 9 AM EDT departure | ±2–3 min |
| **Straight-line fallback** | Haversine × 1.3 road factor, 30 mph avg | ±5–15 min |

Low-quality entries are:
- Shown with gray dashed circles on the map
- Italicized in gray in the Excel export
- Reported separately in statistical tests
- Flagged in the `Data Quality` column of the Excel

### EDT/EST Note

Driving times use Google Maps with a Sunday 9:00 AM **EDT** departure time. During EST months (November–March), the actual departure would be 8:00 AM Eastern, which may produce slightly different traffic estimates (typically ±1–2 minutes). This is an acceptable variance for this analysis.

## Household Weighting

Summary statistics are reported both ways:
- **By address**: each address counts equally (one family of 7 = one single person)
- **By headcount**: weighted by household size (a family of 7 counts 7× more than a single person)

Headcount weighting gives a more accurate picture of actual attendance impact — losing a family of 7 to Dalton matters more than losing one person.

## Known Limitations

1. **K is assumed, not calibrated.** There is no survey or behavioral data to determine the "correct" K. The sensitivity analysis mitigates this by showing which conclusions hold across all reasonable K values.

2. **Driving time is the only factor.** The model does not account for social ties, family/friend networks, pastoral relationships, ministry involvement, worship style preferences, or any non-geographic factor.

3. **Incumbency bias is not applied by default.** The base model treats Jackson and Dalton as interchangeable. The sensitivity analysis shows the impact of adding a switching cost, but no specific bias value is built into the primary outputs.

4. **This analyzes existing members only.** It answers "which current members might prefer Dalton?" — not "is Dalton a good location for a new church?" The latter would require population density data, unchurched population estimates, competing church analysis, and demographic trends.

5. **Straight-line fallback is rough.** The 1.3× road factor and 30 mph assumption may underestimate actual driving times on rural township roads in Stark/Wayne County.

## File Structure

```
New Church Plant/
├── church_map.py                     # Main analysis script (run this)
├── servantkeeper_scraper.py          # Data scraper for ServantKeeper directory
├── FBCJ_Directory #2.xlsx           # Source data (ServantKeeper export)
├── FBCJ_Directory.xlsx              # Original scraper output
├── FBCJ_Directory_Map.csv           # Processed CSV of addresses
├── geocoded_addresses.json          # Cached geocoding + routing results
├── README.md                        # This file
│
│  Generated outputs:
├── FBCJ_Church_Map.html             # Interactive map with boundary lines
├── FBCJ_Church_Map_Time.html        # Driving time contour map
├── FBCJ_Church_Map_Distance.html    # Driving distance contour map
├── FBCJ_Church_Probability_Distance.html  # Distance scatter plot
├── FBCJ_Church_Probability_Time.html      # Time scatter plot
├── FBCJ_Church_Histogram.html       # Time difference distribution
├── FBCJ_Sensitivity_Analysis.html   # K + bias sensitivity chart
├── FBCJ_Analysis_Results.xlsx       # Full Excel export with summary
├── FBCJ_Church_Probability.png      # Static probability image
├── FBCJ_Church_Probability_Distance.png  # Static distance image
└── FBCJ_Church_Probability_Time.png      # Static time image
```

## Configuration

Key parameters in `church_map.py` (top of file):

| Parameter | Default | Description |
|-----------|---------|-------------|
| `JACKSON` | 40.89482, −81.51603 | Jackson church coordinates |
| `DALTON` | 40.79872, −81.69517 | Proposed Dalton location coordinates |
| `K_LOGISTIC_TIME` | 0.15 | Logistic steepness per minute (see above) |
| `ROUTING_ENGINE` | "google" | "google" (requires API key) or "osrm" (free) |
| `GRID_N` | 40 | Boundary grid resolution (40×40 = 1,600 points) |
| `GRID_RADIUS_MI` | 20 | Grid radius from church midpoint |
| `SKIP_STATES` | MX, FL, PA, MI | States to exclude (out of region) |

## Dependencies

```
folium          # Interactive maps
matplotlib      # Contour extraction
numpy           # Numerical operations
openpyxl        # Excel read/write
plotly           # Interactive charts
requests        # API calls (geocoding, routing)
scipy           # KDE, statistical tests, image processing
```
