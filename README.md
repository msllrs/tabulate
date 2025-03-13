# Tabulate for Figma

A Figma plugin that allows you to easily populate table components with structured data. Transform your tables with custom JSON data or generate realistic dummy data with a single click.

## Features

- **Custom Data Import**: Populate your tables with your own structured JSON data
- **Dummy Data Generation**: Instantly fill your tables with realistic, randomized sample data
- **Automatic Cell Detection**: Works with table structures composed of cell components
- **Smart Row Organization**: Automatically organizes cells into rows based on their position
- **Value Layer Targeting**: Updates text layers named "value" within cell components

## Installation

1. Download the plugin files
2. In Figma, go to **Plugins > Development > Import plugin from manifest...**
3. Select the `manifest.json` file from the downloaded files
4. The plugin will be installed and available in your development plugins

## Usage

### Preparing Your Table

This plugin works with tables that are structured as follows:
- Tables composed of cell components
- Cell components containing a text layer named "value"
- Cells organized in rows (either as direct children or within row components)

### Populating with Custom Data

1. Select your table or table rows in Figma
2. Open the plugin
3. Paste your JSON data in the provided text area
4. Click the "Tabulate" button
5. Your table will be populated with the provided data

### Using Dummy Data

1. Select your table or table rows in Figma
2. Open the plugin
3. Click the "Generate & Tabulate" button in the Dummy Data section
4. Your table will be instantly populated with realistic dummy data

## Data Format

The plugin expects data in the following JSON format:

```json
[
  {
    "Name": "Mark Darnalds",
    "Department": "Design",
    "Email": "mark.darnalds@company.com",
    "Location": "London",
    "Access Level": "Admin",
    "Status": "Active"
  },
  {
    "Name": "Wendy Kingsley",
    "Department": "Marketing",
    "Email": "wendy.kingsley@company.com",
    "Location": "New York",
    "Access Level": "Editor",
    "Status": "Active"
  }
]
```

The keys in each object will be used as headers, and the values will be used to populate the corresponding cells.

## How It Works

1. The plugin identifies all cell components in your selection
2. It organizes these cells into rows based on their position
3. It extracts the keys from your JSON data to create headers
4. It populates each cell's "value" text layer with the corresponding data
5. For dummy data, it generates realistic, randomized data that matches the format

## Credits

Developed by LFSGD

## Support

If you encounter any issues or have questions, please file an issue on the GitHub repository.

