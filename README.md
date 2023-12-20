# Update Scraper

Update scraper will automatically go to www.catalog.update.microsoft.com, search for the newest patch available and replace the content of ps1 script (update_tool.ps1) to always have the newest possible links that exist for windows os patches.

## Reuirements

python 3.11+

### 3rd Party libraries used for this project

- requests - to get html
- bs4 - BeatifulSoup to parse html and get needed links from the html

### Github Actions

Workflow will be triggered automatically every month from 8th to 15th.
This range of days has been set to always catch 'patch Tuesday' and get the newest links every month.
