# NSE Options Trading Analyzer

A Python desktop application for analyzing NSE options trading data with strike price optimization and IRR calculation.

## Installation

### Requirements
- Python 3.7+
- pandas
- openpyxl
- tkinter

### Install Dependencies
```bash
pip install pandas openpyxl tkinter requests
```

## Usage

Run the application:
```bash
python main.py
```

1. Set margin percentage (default: 15%)
2. Set lot size multiplier (default: 1)
3. Click "Fetch & Analyze Options"
4. Review results in the table
5. Click "Export to Excel" to save data

## Features

- Calculate 52-week price percentile
- Find optimal CE and PE strike prices with configurable margin
- Calculate option premiums
- Calculate IRR (Internal Rate of Return)
- Export analysis to Excel

## Calculations

### Percentile
```
Percentile = ((Current - 52W Low) / (52W High - 52W Low)) × 100
```

### Strike Price with Margin
```
CE Strike = Nearest Strike × (1 - Margin%)
PE Strike = Nearest Strike × (1 + Margin%)
```

Example (15% margin, Spot = 140):
- Nearest Strike: 140
- CE Strike: 140 × 0.85 = 119 → 120 (nearest)
- PE Strike: 140 × 1.15 = 161 → 160 (nearest)

### IRR Calculation
```
IRR = (Premium / Margin Required) × (365 / Days to Expiry) × 100
```

## Output Format

The application displays and exports:
- Symbol, Spot Price, 52W High/Low
- Percentile
- Lot Size
- CE Strike, Premium, IRR
- PE Strike, Premium, IRR

## Configuration

Edit these values in the code to customize:
- `strike_interval=50` - Change strike price intervals
- `days_to_expiry=30` - Change expiry days for IRR
- `margin_required = strike_price * 0.15` - Change margin requirement

## License

MIT License
