# Changelog
All notable changes to this project will be documented in this file.

## (Unreleased) (dd/mm/yyyy)
### Added

### Changed

### Deprecated

### Removed

### Fixed

### Security

### Internal

## v0.2.1 (21/10/2024)
### Changed
- Update VIKTOR to v14.16.1

### Fixed
- Fix parsing of charts without X-data (fall back to index)

## v0.2.0 (26/07/2024)
### Added
- ExcelChartParser

### Deprecated
- ExcelImageParser will be replaced by ExcelChartParser

## v0.1.9 (23/04/2024)
### Changed
- Updated VIKTOR dependency

## v0.1.8 (15/04/2024)
### Fixed
- Allow for multiple traces using the same category data

## v0.1.7 (02/04/2024)
### Fixed
- Add output type
- Add figure type to titles

## v0.1.6 (14/03/2024)
### Fixed
- Catch unnamed figures

## v0.1.5 (06/03/2024)
### Changed
- Exclude output fields with no values

## v0.1.4 (05/03/2024)
### Removed
- Removed redundant dependencies (from build)

## v0.1.3 (05/03/2024)
### Removed
- Removed redundant dependencies

## v0.1.2 (04/03/2024)
### Changed
- Accommodate for sheet without outputs
- Accommodate for category and data for single figure coming from different sheets in excel file

## v0.1.1 (01/03/2024)
### Added
- Check for empty inputs

## v0.1.0 (13/02/2024)
### Added
- Initial publish