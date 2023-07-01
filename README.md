# autoexcel

## Overview
This repository contains a simple web application that processes an Excel file, identifies entries within a specific time range, and summarizes the data by 15-minute intervals. It is written in plain JavaScript and uses the XLSX library for reading Excel files.

## Features
- File input for Excel files.
- Input fields for specifying start and end times.
- Analyzes data in a column named "closed_time".
- Generates a table summarizing the number of items within 15-minute intervals.
- Provides an option to copy the generated table to the clipboard.

## Getting Started

### Prerequisites
- Web Browser (Chrome, Firefox, etc.)
- Basic knowledge of HTML and JavaScript.

### Cloning the Repository
1. Open your terminal/command prompt.
2. Clone this repository by running: 
```
git clone <repository-url>
```

### Including the XLSX Library
Make sure to include the XLSX library in your HTML file which is necessary for reading Excel files.

You can include it via CDN by adding the following script tag in your HTML:
```html
<script src="https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.17.0/xlsx.full.min.js"></script>
```

### Running the Application
Open the HTML file with a web browser.

## How to Use

1. Choose an Excel file by clicking on the file input button.
2. Enter the start and end times in the format `HH:MM` (24-hour format) and make sure `MM` is always 00. 
3. The application will process the file and generate a table below the input fields. This table shows the number of entries within each 15-minute interval between the specified start and end times.
4. You can copy the generated table to the clipboard by using the "Copy to Clipboard" button.

## Contributing
Feel free to submit issues, fork the repository and send pull requests!
