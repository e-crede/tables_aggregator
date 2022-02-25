# tables_aggregator
Backup various Excel tables &amp; aggregate them into a database

# ClusterField

## Overview

Business Intelligence application that periodically copies XLSX or CSV tables into a local directory, stores them by date then copies the information into an sqlite database for easier aggregation and analysis

## Use Cases

Excel keeps a lot of businesses running, therefore within organisations there tends to be a lot of excel tables stored in different locations, to cater for different users different use cases.

The information itself tends to be related and potentially even have common keys (e.g. project or task ID’s).

Thus this tool aims to centralise this information and save it into a single database, which can then be queried directly by the user or some application.

Thus such a quick query would provide a quick overview to make more informed daily business decisions

## Features

### File Transfer

- Intended to be run as a Scheduled daily task (cron job)
- Comparison of latest grabbed and previous files
    - [ ] Store only unique files
    - [ ] Compare by edit date
    - [ ] Compare by checksum
- [ ] Organise stored files by date - a single file per-day
- [ ] Files that have to be retrieved are configured within a configuration file
- [ ] File names matching
    - [ ] must have dynamic pattern matching support (e.g. in case source files change their name on daily / weekly basis)
    - [ ] must have ability to configure conditions - e.g. grab only newest file in a directory

## Database

- [ ] Data Stored in an SQLite database
- [ ] Table Headers automatically recognised
- [ ] Configuration on  which rows and columns to import from a file

## Exclusions and Limitations

- File Transfer
    - No cleanup of files or file storage monitoring
- Database
    - No archiving or cleanup functions
    - Primary keys across different excel sheets will differ
