# BorsdataDownloader

A program that opens a web browser with Selenium, logs into [Börsdata](www.borsdata.se) and downloads Excel-files with data for a specific strategy. The Excel-file is processed with Pandas, sorted according to relevant financial ratios and archived for potential backtesting in the future.

The currently implemented strategies are Magic Formula and Acquirer's Multiple.

Requires a Börsdata Premium subscribtion with a user_credentials file.

Logs into Börsdata on its own.

<img src="https://github.com/hataloo/BorsdataDownloader/blob/master/BorsdataShowcase/Login.png" width="500">

Selects a strategy, fills in desired markets and industries.

<img src="https://github.com/hataloo/BorsdataDownloader/blob/master/BorsdataShowcase/SelectStrat.png" width="280">
<img src="https://github.com/hataloo/BorsdataDownloader/blob/master/BorsdataShowcase/SelectMarket.png" width="280">
<img src="https://github.com/hataloo/BorsdataDownloader/blob/master/BorsdataShowcase/SelectIndustry.png" width="280">

Uses Börsdata's export feature to download the data in an Excel-file.

<img src="https://github.com/hataloo/BorsdataDownloader/blob/master/BorsdataShowcase/Downloaded.png" width="400"><img src="https://github.com/hataloo/BorsdataDownloader/blob/master/BorsdataShowcase/ScrapedExcel.png" width="400">

The Excel-file is processed and filtered according to the selected strategy. For example, Magic Formula filter companies with a market cap below 500 million SEK, rank companies separetely in ROC (Return On Capital) and EBIT (Earnings Before Interest and Taxes) and combines the ranks into the Magic rank.

<img src="https://github.com/hataloo/BorsdataDownloader/blob/master/BorsdataShowcase/ProcessedExcel.png" width="425"> <img src="https://github.com/hataloo/BorsdataDownloader/blob/master/BorsdataShowcase/DataframeSpyder.png" width="425">
