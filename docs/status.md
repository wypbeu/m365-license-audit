---
title: Status
date: 2026-04-17
tags: [status, log, meta]
---

# Status

Chronological session log. Newest entry at top. Written at session end via `/session-close` — one entry per meaningful session, covering what changed, decisions made, and what's next.

---

## 2026-04-17

**PS5 compatibility fix across all three scripts.** Nathan (Ramsay) hit a `ParserError: Unexpected token '??'` running `Get-LicenceInventory.ps1` on Windows PowerShell 5.x. The `??` null-coalescing operator requires PS7+. Replaced all seven occurrences across `Get-LicenceInventory.ps1`, `Get-UserLicenceMap.ps1`, and `Find-LicenceWaste.ps1` with `if ($null -ne ...) { ... } else { ... }` equivalents. Committed and pushed as `927862d`.

**Next**: confirm Nathan can run the full audit chain successfully on his machine; check whether any other PS7-specific syntax (e.g. ternary `? :`) is used elsewhere in the scripts.

---
