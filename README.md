# Things to Todoist Converter

A Python script that uses AppleScript to extract data from Things 3 and convert it to Todoist CSV format. It preserves projects, areas, tasks, tags, due dates, and recurring task schedules.

## Requirements

- macOS (for AppleScript support)
- Python 3.6+
- Things 3 app
- pandas (optional, for better CSV handling)

## Installation

1. Clone this repository:
   ```bash
   git clone https://github.com/yourusername/things-to-todoist-converter.git
   cd things-to-todoist-converter
   ```

2. (Optional) Install pandas for better CSV handling:
   ```bash
   pip install pandas
   ```

## Usage

1. Make sure Things 3 app is running
2. Run the script:
   ```bash
   python things_applescript_to_todoist.py output.csv
   ```

### Command-line Options

- `--skip-completed`: Skip completed tasks
- `--areas-only`: Only include areas
- `--projects-only`: Only include projects
- `--todos-only`: Only include to-dos

### Examples

```bash
# Export everything
python things_applescript_to_todoist.py output.csv

# Skip completed tasks
python things_applescript_to_todoist.py output.csv --skip-completed

# Only export areas
python things_applescript_to_todoist.py output.csv --areas-only

# Only export projects
python things_applescript_to_todoist.py output.csv --projects-only

# Only export to-dos
python things_applescript_to_todoist.py output.csv --todos-only
```

## Import to Todoist

1. Go to Todoist Settings > Import/Export
2. Choose 'Import from CSV file'
3. Select the generated CSV file

## Features

- Extracts areas, projects, and to-dos from Things 3
- Preserves hierarchy (areas > projects > tasks)
- Converts Things tags to Todoist tags
- Handles recurring tasks
- Preserves due dates
- Converts checklist items to subtasks
- Marks completed items

## How It Works

The script uses AppleScript to communicate with Things 3 and extract data. It then processes this data and converts it to Todoist's CSV format. The script handles special cases like missing areas, completed tasks, and provides flexible options for what to include in the export.

## License

MIT 