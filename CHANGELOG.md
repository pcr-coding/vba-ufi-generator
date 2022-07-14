# VBA UFI Generator - Changelog
**0.1 (2021-03-08)**  
Initial commit

**0.2 (2021-03-09)**  
[fix] Added support for spanish VATs. Samples were missing in the official manual table B-1, so I created a own (verified with the official online generator).  
[add] Licence information GPL-3.0-only.

**0.3 (2022-07-14)**
[fix] Link in Readme to new UFI Developers Manual to version 1.5 (2022-01).
[change] Implement changes from UFI Developers Manual version 1.4 (2021-11) and 1.5 (2022-01)
[change] VAL001 includes check for hyphens and format structure now, incl. adjusted error reporting (according dev-man 1.4).
[add] Support for Northern Ireland XN (according dev-man 1.5).
[add] Unit test for GB|XN.
[fix] EL was not accepted for GR, now implemented to accept both GR and EL.
[add] Unit test for GR|EL.
