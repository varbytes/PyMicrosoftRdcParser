# PyMicrosoftRdcParser.py

Simple script to parse macOS application 'Microsoft Remote Desktop' for bookmarked RDP connections and export to Excel.

Application database file location: `~\Library\Containers\com.microsoft.rdc.macos\Data\Library\Application Support\com.microsoft.rdc.macos\com.microsoft.rdc.application-data.sqlite`

## Usage

```
PyMicrosoftRdcParser.py [-h] [--db DB] [--outfile OUTFILE]
```

Options:
```
  -h, --help         show this help message and exit
  --db DB            Path to the SQLite database file.
  --outfile OUTFILE  Path to output Excel file.
```
