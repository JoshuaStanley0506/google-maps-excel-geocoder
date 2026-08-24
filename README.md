# Google Maps Excel Geocoder

A small production-oriented Python utility that reads Google Maps links from an Excel workbook, follows redirects, extracts geographic coordinates, converts them to DMS notation, and writes the results to a dedicated worksheet.

## Features

- Handles Google Maps short links and redirected URLs.
- Prioritizes place coordinates from `!3dLAT!4dLON`.
- Supports `q=`, `ll=`, and `@LAT,LON` coordinate formats as fallbacks.
- Writes decimal and DMS latitude/longitude.
- Preserves the source workbook and adds a `Map Coordinates` sheet to a new output workbook.
- Shows a progress bar for larger workbooks.
- Logs request failures, missing coordinates, malformed rows, and summary statistics.
- Includes unit tests for the coordinate parsing and validation logic.
- Supports both a command-line input path and a graphical file picker.

## Expected input columns

The first worksheet must contain these columns:

- `Head of Family`
- `Contact Address`
- `Prayer Group`
- `Maps Link`

Additional columns are allowed and are ignored by the coordinate output sheet.

## Installation

Python 3.10+ is recommended.

```bash
git clone https://github.com/JoshuaStanley0506/google-maps-excel-geocoder.git
cd google-maps-excel-geocoder
python3 -m venv .venv
source .venv/bin/activate
python -m pip install --upgrade pip
pip install -r requirements.txt
```

## Usage

### Graphical mode

Run without an input filename to open a file picker:

```bash
python mapmylink.py
```

### Command-line mode

```bash
python mapmylink.py input.xlsx
```

Specify an output path when required:

```bash
python mapmylink.py input.xlsx --output geocoded.xlsx
```

If no output path is supplied, the application writes `mapmylink.xlsx` beside the input workbook.

## Output

The generated workbook contains a `Map Coordinates` sheet with:

| Column | Description |
|---|---|
| Head of Family | Copied from input |
| Contact Address | Copied from input |
| Prayer Group | Copied from input |
| Maps Link | Original link |
| Expanded Maps Link | Final URL after redirects |
| Decimal Latitude | Latitude in decimal degrees |
| Decimal Longitude | Longitude in decimal degrees |
| DMS Latitude | Degrees/minutes/seconds latitude |
| DMS Longitude | Degrees/minutes/seconds longitude |

## Coordinate extraction strategy

The parser deliberately prefers coordinates that identify the actual place rather than map viewport coordinates:

1. `!3dLAT!4dLON` — place coordinates
2. `q=LAT,LON` — query coordinates
3. `ll=LAT,LON` — `ll` coordinates
4. `@LAT,LON` — viewport fallback

Viewport coordinates can represent the center of a map rather than the business/place itself, so the application logs a warning when that fallback is used.

## Testing

Run the test suite with:

```bash
pytest -q
```

## Logging

Runtime logs are written to `mapmylink.log`. The log contains per-row failures and a final processing summary.

Do not commit input workbooks, output workbooks, or log files containing sensitive data. The repository's `.gitignore` excludes these artifacts by default.

## Production notes

- The tool does not require a Google Maps API key because it follows the supplied web URLs and parses the resulting URL.
- Google Maps URL behavior can change. Keep the extraction tests updated if Google changes its URL formats.
- Network access is required for redirect expansion.
- Large batches should be run with reasonable network limits and reviewed for failed rows.

## License

MIT. See `LICENSE`.
