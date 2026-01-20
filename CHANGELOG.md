#####

# Changelog

All notable changes to "files-au3" will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

Go to [legend](#legend-types-of-changes) for further information about the types of changes.

## [Unreleased]

## [0.3.0] - 2026-01-20

### Added

- Community standard 'Code of conduct' document. [7e39247](https://github.com/AutoIt-Community/files-au3/commit/7e39247a9183df7921d9e60d66117ae45d8eacf1)
- Community standard 'SECURITY.md' file. [f706504](https://github.com/AutoIt-Community/files-au3/commit/f706504ddaf480f73c7ce89633b377d620c1a421)
- Initial drag and drop support (no file movement yet). [c6aa277](https://github.com/AutoIt-Community/files-au3/commit/c6aa277cf1032c9edd9e4e11792782cdb4894993)

### Documented

- Update README.md file. [0c2ded0](https://github.com/AutoIt-Community/files-au3/commit/0c2ded05f10b873f3539e37e079591b367d12e09)
- Update README to display contributors. [c7755a4](https://github.com/AutoIt-Community/files-au3/commit/c7755a40988cd7b1fcfaa81aaaf894ecc02f56bf)
- Add contribution instructions via CONTRIBUTING.md file. [8273d4a](https://github.com/AutoIt-Community/files-au3/commit/8273d4ab2d63a365db8a88f41f1b11e3c3ca0659)
- Add link to contribution section in README.md file. [d7a1532](https://github.com/AutoIt-Community/files-au3/commit/d7a153200c3c6a0c0f3bf8c4121a0356b241b8c4)
- Update README by license note and third-party author credits. [1455cdd](https://github.com/AutoIt-Community/files-au3/commit/1455cdd9fb132fc3d2dc4e6ca28c68436816f358)
- Update LICENSE.md file. [1bf48dc](https://github.com/AutoIt-Community/files-au3/commit/1bf48dc07411b89dc94ffa8c744774bd1337a0be)

### Fixed

- Resizing issue with child GUI frames and controls from GUIFrame UDF. [c6aa277](https://github.com/AutoIt-Community/files-au3/commit/c6aa277cf1032c9edd9e4e11792782cdb4894993)
- Fixed issue in dark mode where separator frame flickered during initial GUI startup. [d1b99fa](https://github.com/AutoIt-Community/files-au3/commit/d1b99fa6cf3ca036e56c9e105d22e823655ecd70)

### Thanks to the contributors (committers)

[@sven-seyfert](https://github.com/sven-seyfert), [@WildByDesign](https://github.com/WildByDesign)

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

### Thanks to the contributors (committers)

[@DonChunior](https://github.com/DonChunior), [@WildByDesign](https://github.com/WildByDesign)

## [0.1.0] - 2026-01-13

### Added

- Initial commit (code base, assets, GitHub files). [17aab44](https://github.com/AutoIt-Community/files-au3/commit/17aab4409ea283ad181aed28984cb42c2cb8591f)

### Thanks to the contributors (committers)

[@sven-seyfert](https://github.com/sven-seyfert), [@WildByDesign](https://github.com/WildByDesign)

[Unreleased]: https://github.com/AutoIt-Community/files-au3/compare/v0.3.0...HEAD
[0.3.0]: https://github.com/AutoIt-Community/files-au3/compare/v0.2.0...v0.3.0
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
