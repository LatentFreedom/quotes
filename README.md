# Quote Scraper
Srape **[www.brainyquote.com](www.brainyquote.com)** for quotes and fill an excel sheet given a keyword.

## Firefox
`geckodriver` required in path before running. You can find geckodriver at **[https://github.com/mozilla/geckodriver/releases](https://github.com/mozilla/geckodriver/releases)**

## Excel
Create an excel file with the name `quotes.xls`. There will be a new worksheet created for every new word that is queried.

## Usage
```
python quotes.py <term>
```