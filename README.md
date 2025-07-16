# 10K_downloader
> Downloads most recent annual report 10K filing from the SEC's Edgar database.
> This program needs 2 inputs, an Excel sheet (containing tickers and company names) and a txt containing a JSON string mapping SEC CIK numbers to tickers. An example CIK dictionary is provided in this repo.
> Output is written in txt form for frequency analysis. To make it more legible to humans consider saving filing as docx.
> Debugging: the module selenium is in the experimental phase. Any bugs are best solved by running program from the site-packages directory.
