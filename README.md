# xlDiff

**A more user-friendly, intuitive, efficient, and resource-prioritized spreadsheet comparison application** ✨

Keywords: **xls compare xlsx compare diff compare excel sheet Spreadsheet**

![Preview](https://github.com/orunco/xldiff-release/blob/main/preview.gif)

You have two Excel workbooks, or two versions of the same workbook, and you want to see the changes between them. Alternatively, you may want to identify potential issues. Simple comparisons are always straightforward, such as:

- General cell value changes
- Sheet name changes
- And so on

There are plenty of existing tools to support this, or you can write code to perform the comparison yourself. However, in certain special scenarios, things can get a bit tricky:

- **Entire column moves, copies, or pastes**: Data from an entire column is relocated to another position, even though the actual content may not have changed.
- **Understandability of comparison results**: The comparison results between two workbooks can become so complex that they are difficult to interpret. In such cases, improving the clarity of the comparison report interface is the final critical step in resolving comparison issues.
- **Accuracy of formula values**: The foundation of comparison lies in accurately parsing the values in cells, including formula values and formatted values. If the results are obtained directly from Microsoft Office's native Excel, they are accurate, but performance issues can be challenging to resolve for extremely large spreadsheets. If third-party components are used for direct parsing, the results may not match those of Microsoft Excel.
- **Extra-large Excel files**: Comparing small datasets is always easy. However, the xlsx format supports up to 1,048,576 rows × 16,384 columns. For medium to large datasets, it is necessary to balance the performance of reading xlsx files, comparison algorithms, and memory resource requirements.

xlDiff aims to address these issues to create a more user-friendly, intuitive, and efficient spreadsheet comparison application.

## Usage Steps

```markdown
Download the `Orunco.xlDiff-win-x64.zip` version locally
Unzip and run `xldiff_desktop-Native-win-x64.exe`
Select files and configure which Sheets to compare, the header for each Sheet, and which fields to compare
Click "Compare"
View the results
```

## Core Concepts

Unlike most existing tools, xlDiff requires Sheet pairing and column header pairing before comparison. Of course, xlDiff will automatically initialize these configurations based on certain rules, but they can also be manually configured as needed. Sample data can be loaded from the Sample folder in the installation directory.

## Feature Overview

Key steps for comparison:

1. Select the two files to compare
2. Perform Sheet pairing: For example, compare Sheet1 with Sheet2
3. Set the header information for each Sheet: For example, specify that the header for Sheet1 is in the second row, allowing xlDiff to read its column names
4. Perform column pairing within the Sheet: For example, compare Column A with Column B
5. Execute the comparison algorithm and analyze the results

## Data Privacy and Security

We respect your privacy and take great care to protect your information. We do not upload any local user data or data involving personal information. The original xls/xlsx files are read-only, with no write or modification operations performed on them, ensuring no damage to your local data.

## LICENSE

Except for third-party components, all rights are reserved by Orunco.

## Disclaimer

The content of this document has been carefully created and arranged. We assume no responsibility for the accuracy, completeness, or timeliness of the provided content. While we strive to ensure the accuracy of the final comparison results, we cannot fully guarantee it.

## Antivirus Software

We have not purchased third-party digital certificates, so antivirus software may occasionally flag false positives.

Before releasing any version, we conduct online virus scans to confirm that no viruses are present.

## Changelog
### 25.07.0.0

- Added support for Simplified Chinese, including the xlDiff interface, Spreadsheet data content, and numeric formats
- Added support for reading xls and xlsx files
- Added support for Sheet pairing operations
  - Added support for clearing all Sheets at once
  - Added support for selecting individual Sheets for comparison
- Added support for setting header information for Sheets: The right side of the current interface is only for data preview and header configuration, displaying up to the first 1,000 rows (this does not mean the data is unread)
- Added support for column pairing within Sheets
  - Added support for re-auto-pairing
  - Added support for removing partial field pairings
  - Added support for pairing by index rather than name
- Right-click menu now supports column ignore settings
- Maximized support for the accuracy of formatted cell values in a Simplified Chinese environment
- Formula values are not recalculated; the values saved in the file at the time of reading are used to ensure the highest possible accuracy of formula calculations
- Added support for a quick comparison algorithm: If the difference between rows is below a certain threshold, it is marked as a modified row; otherwise, it is marked as a deleted/added row
- Added support for exporting comparison results (requires an application capable of handling xls/xlsx files on the local machine). Users can reconstruct the original data and differences to the greatest extent based on the results.
- Only supports comparison of cell values
- Significant efforts have been made to optimize performance and memory usage
- Only supports Windows 10 64-bit and Windows 11 64-bit environments