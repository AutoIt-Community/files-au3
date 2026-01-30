#####

# Changelog

All notable changes to "files-au3" will be documented in this file.

The format is based on [Keep a Changelog](https://keepachangelog.com/en/1.1.0/),
and this project adheres to [Semantic Versioning](https://semver.org/spec/v2.0.0.html).

Go to [legend](#legend-types-of-changes) for further information about the types of changes.

## [Unreleased]

### Changed

- Clear previous statusbar item count. [b34bb50](https://github.com/AutoIt-Community/files-au3/commit/b34bb50728fdda8357e91317583d08c4c7a56952)
- Apply code review suggestions and add GUI cursor override for slow loading directories. [6c17901](https://github.com/AutoIt-Community/files-au3/commit/6c17901ecf7efe12626ded122e400835a466b622)

### Documented

- Update license file. [1762860](https://github.com/AutoIt-Community/files-au3/commit/1762860ddd44aaec0627327e8c5e201c8db2901d)

### Fixed

- Enumeration of items on statusbar by ensuring that ListView update is complete. [38e9979](https://github.com/AutoIt-Community/files-au3/commit/38e99791d29cabfb2558a94ac9e1f3603810b0de)

### Refactored

- Apply alphabetical order for include statements. [a8f4655](https://github.com/AutoIt-Community/files-au3/commit/a8f4655690ff5ad7b7b559556465265264eab2be)
- Use local scope of a variable instead of global (because not needed to be global). [49f153d](https://github.com/AutoIt-Community/files-au3/commit/49f153d6a13ebd13be7efcd0fb3765f0fbcc548b)
- Use one-time used constants directly in function instead of global constant. [0a0a8aa](https://github.com/AutoIt-Community/files-au3/commit/0a0a8aaa893ff7be9c2b79e7c9414f5868c122ca)
- Syntax clean-up by tidy usage. [70b275b](https://github.com/AutoIt-Community/files-au3/commit/70b275b8a09039d6652728452286256f6449e621)
- Drag and drop improvements. [195e421](https://github.com/AutoIt-Community/files-au3/commit/195e421ae2b2aeb26e2112518922a58e05ddfc23)
- Apply return early pattern (code review suggestion). [68c01dd](https://github.com/AutoIt-Community/files-au3/commit/68c01dd07bc2c30527feeeb425f1ed3f947742ff)

### Removed

- Get rid of unused variables. [6ab1a5b](https://github.com/AutoIt-Community/files-au3/commit/6ab1a5b02fe7aba57fd5243ae18a8f83ad2903c5)
- Unused WinAPI function (including return value) in _InitDarkSizebox. [c225e09](https://github.com/AutoIt-Community/files-au3/commit/c225e0968a0d81f09e8238a654f52b694e5f3a1b)
- Dead code (unused WinAPI function call). [c2127d3](https://github.com/AutoIt-Community/files-au3/commit/c2127d3ebd6e4c719bbf68ab003540fd042898bf)

### Styled

- Trival adjustments to comments (syntax). [a65308f](https://github.com/AutoIt-Community/files-au3/commit/a65308fee10bcc79b0d804ce03beece9bff43f93)

### Thanks to the contributors (committers)

[@sven-seyfert](https://github.com/sven-seyfert), [@WildByDesign](https://github.com/WildByDesign)

## [0.4.0] - 2026-01-22

### Added

- Tooltips for drag and drop functionality (no file movement yet). [589c563](https://github.com/AutoIt-Community/files-au3/commit/589c5638c482e5c339570106caec132cb6c5d6fa)
- Community standard files like issue templates and pull request template. [50c78a2](https://github.com/AutoIt-Community/files-au3/commit/50c78a2c9c5c3a87f3c5d474f1a88b52456f7d4e)

### Changed

- Some of the drag and drop related structure. [589c563](https://github.com/AutoIt-Community/files-au3/commit/589c5638c482e5c339570106caec132cb6c5d6fa)

### Fixed

- Some issues with drag and drop tooltips. [3c4c977](https://github.com/AutoIt-Community/files-au3/commit/3c4c97779b72303953594ad6ad7cebd99a7218d4)

### Refactored

- Cleanup of dead code. [07a6e82](https://github.com/AutoIt-Community/files-au3/commit/07a6e824e8f296e2272e9893c86b2b8145fc27d0)
- Unused code and general code cleanup. [c95ad59](https://github.com/AutoIt-Community/files-au3/commit/c95ad59680309bc57a200fa55ef4300f25b24422)

### Thanks to the contributors (committers)

[@sven-seyfert](https://github.com/sven-seyfert), [@WildByDesign](https://github.com/WildByDesign)

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

[Unreleased]: https://github.com/AutoIt-Community/files-au3/compare/v0.4.0...HEAD
[0.4.0]: https://github.com/AutoIt-Community/files-au3/compare/v0.3.0...v0.4.0
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
