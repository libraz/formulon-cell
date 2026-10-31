# Security policy

## Reporting a vulnerability

Report privately by email to `libraz@libraz.net`, not through the public issue tracker.

Include the workbook or the steps that trigger it, what happened, the affected package and version, the browser, and a minimal reproduction if you have one. Expect an acknowledgement within a few days.

## Supported versions

Pre-1.0. Fixes land on the latest release only; older versions are not patched.

## What is in scope

formulon-cell renders workbooks it did not write and takes input from the clipboard and from the user, inside the host's page. In scope:

- A workbook, a pasted value or typed input whose content runs as script in the host page: cell text, comments, sheet names, chart titles, hyperlinks, shapes or any other workbook content reaching the DOM as markup.
- A hyperlink that opens a `javascript:` URL or another scheme the spreadsheet does not present as a link, or that navigates the host page without the user following it.
- Anything that lets workbook content read from or write to the host page outside the spreadsheet's own element, its store, or the files the user chose to open and save.
- An interaction policy (`viewerPolicy`, `fixedFormPolicy` or a custom policy) that a user can get around through any input path to change a cell the policy protects.

## What is not in scope

- Wrong formula results, crashes or hangs inside the calc engine. Report those to the [formulon](https://github.com/libraz/formulon) project, which has its own security policy.
- Behaviour that differs from a desktop spreadsheet application. Report it as a normal issue.
- Sheet and workbook protection as stored in the file. It controls editing in the spreadsheet UI and is not a security boundary.
- Findings that require an attacker to already control the host page or the code that mounts the spreadsheet.
