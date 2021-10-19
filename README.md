# Find URLs in a CDX index

Takes a spreadsheet of URLs and checks if each unique URL exists in a CDX index.
Outputs a new spreadsheet of the URLs that do exist in the the CDX index along with their three latest occurrences.

Currently assumes URLs will be in column E of the input spreadsheet.


## Installation
Use the package manager [pip](https://pip.pypa.io/en/stable/) to install necessary packages. 

```bash
pip install openpyxl requests datetime
```

## Usage

Takes three arguments:

1. Path to the input spreadsheet
2. Location to save the output spreadsheet
3. URL of the CDX index

```bash
python3 find_urls_in_cdx_index.py <path to input spreadsheet> <path to destination folder> <url of cdx index>
```