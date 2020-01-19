# BorsdataDownloader

A program that opens a web browser with Selenium, logs into [Börsdata](www.borsdata.se) and downloads Excel-files with data for a specific strategy. The Excel-file is processed with Pandas, sorted according to relevant financial ratios and archived for potential backtesting in the future.

The currently implemented strategies are Magic Formula and Acquirer's Multiple.

Requires a Börsdata Premium subscribtion with a user_credentials file.

![Image](https://github.com/hataloo/BorsdataDownloader/blob/master/BorsdataShowcase/Login.png =200x)
