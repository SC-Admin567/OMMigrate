# License

## MIT License

**Copyright (c) 2026 Kirk Shallcross - Shallcross Consulting**

Permission is hereby granted, free of charge, to any person obtaining a copy
of this software and associated documentation files (the "Software"), to deal
in the Software without restriction, including without limitation the rights
to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
copies of the Software, and to permit persons to whom the Software is
furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all
copies or substantial portions of the Software.

**THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
SOFTWARE.**

---

## What This Means in Plain English

You are free to:

- **Use** OMMigrate personally, professionally, or commercially
- **Copy** and distribute OMMigrate to others
- **Modify** OMMigrate to suit your environment or client needs
- **Build** commercial products or services on top of OMMigrate
- **Include** OMMigrate in paid IT consulting engagements

You must:

- **Keep the copyright notice** — the credit block naming Kirk Shallcross - Shallcross Consulting
  as Originator & Architect must remain intact in any copy or derivative work
- **Include this license** — any distribution of OMMigrate or a substantial
  portion of it must include this LICENSE.md file

You cannot:

- **Hold the author liable** — OMMigrate is provided as-is with no warranty.
  Test thoroughly in a non-production environment before running against
  live Outlook profiles. Always verify backups before proceeding.

---

## Attribution

If you distribute OMMigrate or a derivative work, the following attribution
must appear in the documentation, credits section, or about screen:

```
OutlookMailMigrator (OMMigrate)
Originator & Architect:    Kirk Shallcross - Shallcross Consulting
Implementation Specialist: Anthropic Claude AI
https://github.com/SC-Admin567/OMMigrate
```

---

## Third-Party Components

OMMigrate uses only:

- **Windows PowerShell 5.1** — included with Windows 10 and Windows 11
- **Microsoft Outlook COM Object Model** — part of Microsoft Office
- **Windows Registry APIs** — part of the Windows operating system

No third-party libraries, NuGet packages, or external dependencies are
included or required. OMMigrate is entirely self-contained.

---

## Disclaimer

OMMigrate interacts with Microsoft Outlook via the documented and publicly
supported Outlook COM Object Model (OOM). It reads Windows Registry keys
that contain the operator's own configuration data. It does not reverse
engineer, circumvent, or violate any Microsoft End User License Agreement.

Use of OMMigrate is at the operator's own risk. Always:

1. Verify PST backups exist and are valid before proceeding past Script 01
2. Test in a non-production Outlook profile before running against live accounts
3. Keep backup PST files for a minimum of 30 days after migration completes
4. Verify Outlook Rules are routing email correctly after Script 03 completes
5. Understand that OMMigrate is a one-way migration tool -- it does not provide
   automated rollback from IMAP back to POP3. Manual recovery procedures are
   documented in README.md under "POP3 Manual Recovery"

---

```
Version: 1.5.2
OutlookMailMigrator (OMMigrate) v1.5.2
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
Originator & Architect:    Kirk Shallcross - Shallcross Consulting
Implementation Specialist: Anthropic Claude AI
Inception Date:            May 2026
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
"Automating the Outlook migration Google suggested couldn't be automated."
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
```
