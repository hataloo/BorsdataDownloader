# BorsdataDownloader

A program that opens a web browser with Selenium, logs into [Börsdata](www.borsdata.se) and downloads Excel-files with data for a specific strategy. The Excel-file is processed with Pandas, sorted according to relevant financial ratios and archived for potential backtesting in the future.

The currently implemented strategies are Magic Formula and Acquirer's Multiple.

Requires a Börsdata Premium subscribtion with a user_credentials file.

Logs into Börsdata on its own.

<img src="https://github.com/hataloo/BorsdataDownloader/blob/master/BorsdataShowcase/Login.png" width="500">

Selects a strategy, fills in desired markets and industries.

<img src="https://github.com/hataloo/BorsdataDownloader/blob/master/BorsdataShowcase/SelectStrat.png" width="500">
<img src="https://github.com/hataloo/BorsdataDownloader/blob/master/BorsdataShowcase/SelectMarket.png" width="500">
<img src="https://github.com/hataloo/BorsdataDownloader/blob/master/BorsdataShowcase/SelectIndustry.png" width="500">

Uses Börsdata's export feature to get the data in an Excel-file.

<img src="https://github.com/hataloo/BorsdataDownloader/blob/master/BorsdataShowcase/Downloaded.png" width="500">
