# Security Policy

## Supported versions

Security fixes are provided for the latest published release.

| Version | Supported |
| --- | --- |
| 1.13.x | Yes |
| Earlier versions | No |

## Reporting a vulnerability

Do not open a public issue for a vulnerability that could expose reading data,
local paths, credentials, device identifiers or arbitrary file contents.

Use GitHub's private **Report a vulnerability** flow in the repository Security
tab. Include:

- the affected version and executable edition;
- a concise description of the impact;
- reproduction steps using synthetic data;
- any suggested mitigation.

Never include a real Kindle sidecar, Calibre database, account identifier,
credential, token or personal reading record. The maintainer will acknowledge a
complete report and coordinate disclosure after a fix is available.

## Data-handling boundary

The application is designed to read a mounted Kindle and write only to the
configured output directory on the computer. It does not require cloud
credentials. Reports and cache files can contain book titles, authors, paths and
reading history; treat the output directory as private user data.
