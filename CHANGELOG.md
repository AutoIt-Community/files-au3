#####

# Changelog

All notable changes to "files-au3" will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

Go to [legend](#legend-types-of-changes) for further information about the types of changes.

## [Unreleased]

### Added

- Added initial drag and drop structure (no file movement yet)

## [0.2.0] - 2026-01-14

### Added

- Painting over white menubar line in dark mode. [9297bd6](https://github.com/AutoIt-Community/files-au3/commit/9297bd6755d656c5d45dd5d86013b97569603342)

### Changed

- Apply OnEvent mode for lower idle CPU usage and faster GUI responsiveness. [9297bd6](https://github.com/AutoIt-Community/files-au3/commit/9297bd6755d656c5d45dd5d86013b97569603342)

### Fixed

- Missing variable declaration added. [f4d1734](https://github.com/AutoIt-Community/files-au3/commit/f4d173489b84bf5f4d9ebcb9e1bc8ade0020a932)
- Unreachable code in ApplyDPI function. [604f398](https://github.com/AutoIt-Community/files-au3/commit/604f398a5d89fbbf06fe6bf1dff2a5a4739f3839)

### Removed

- Duplicate variable declarations. [15cb1e2](https://github.com/AutoIt-Community/files-au3/commit/15cb1e213603c0d849d24a3c4aa873ebb80e2047)
- Duplicate code for button refresh when changing themes. [0e3fa72](https://github.com/AutoIt-Community/files-au3/commit/0e3fa72a8325cd0bfdc843b120c0cbad0b463331)
- Unused code with logic error. [bd0ef15](https://github.com/AutoIt-Community/files-au3/commit/bd0ef1534f731f86a4f97d41f4109fb0deacde83)
- Removed usage of SetSysColors. [9297bd6](https://github.com/AutoIt-Community/files-au3/commit/9297bd6755d656c5d45dd5d86013b97569603342)

### Contributor Acknowledgment

Thanks to:
- [@DonChunior](https://github.com/DonChunior)
- [@WildByDesign](https://github.com/WildByDesign)

## [0.1.0] - 2026-01-13

### Added

- Initial commit (code base, assets, GitHub files). [17aab44](https://github.com/AutoIt-Community/files-au3/commit/17aab4409ea283ad181aed28984cb42c2cb8591f)

### Contributor Acknowledgment

Thanks to:
- [@sven-seyfert](https://github.com/sven-seyfert)
- [@WildByDesign](https://github.com/WildByDesign)

[Unreleased]: https://github.com/AutoIt-Community/files-au3/compare/v0.2.0...HEAD
[0.2.0]: https://github.com/AutoIt-Community/files-au3/compare/v0.1.0...v0.2.0
[0.1.0]: https://github.com/AutoIt-Community/files-au3/releases/tag/v0.1.0

---

### Legend - Types of changes

- `Added` for new features.
- `Changed` for changes in existing functionality.
- `Deprecated` for soon-to-be removed features.
- `Documented` for documentation only changes.
- `Fixed` for any bug fixes.
- `Refactored` for changes that neither fixes a bug nor adds a feature.
- `Removed` for now removed features.
- `Security` in case of vulnerabilities.
- `Styled` for changes like whitespaces, formatting, missing semicolons etc.

##

[To the top](#changelog)
