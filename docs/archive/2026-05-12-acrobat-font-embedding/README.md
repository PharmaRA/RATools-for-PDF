# Acrobat Font Embedding Investigation Archive

Date: 2026-05-12
Project: `RATools-for-PDF`
Topic: Investigating whether Adobe Acrobat Pro can be automated to run Preflight fixup `Embed missing fonts`

## Goal

Archive the current investigation state for the suspended `embed_nonstandard_fonts` feature.

The target capability was to get as close as possible to Acrobat Pro's:

`Print Production -> Preflight -> Acrobat Pro DC 2015 Profiles -> PDF fixups -> Embed missing fonts`

## Outcome Summary

The investigation confirmed that Acrobat Pro contains the exact fixup we want, but the currently discovered COM automation surface does not expose a direct way to execute it.

## What Was Verified

### 1. Acrobat COM automation works at a basic level

Confirmed through `win32com`:

- `AcroExch.App`
- `AcroExch.AVDoc`
- `AcroExch.PDDoc`

The probe can:

- launch Acrobat
- open a real PDF
- obtain a live `PDDoc`
- query page count and file name

### 2. `PrintProduction` is reachable from COM menu automation

Observed:

- `MenuItemIsEnabled("PrintProduction") == true`
- `MenuItemExecute("PrintProduction") == true`

But this did not produce an observable transition in:

- `GetActiveTool`
- `GetNumAVDocs`
- `GetActiveDoc`

So it appears to open or signal the feature area, but not directly execute the desired Preflight fixup.

### 3. Direct `Preflight` menu invocation was not exposed

Observed:

- `MenuItemIsEnabled("Preflight") == false`
- `MenuItemExecute("Preflight")` is not a usable path through the current probe

This suggests the user-facing Preflight panel is not mapped to a simple top-level menu command name that `AcroExch.App` can play back directly.

### 4. JavaScript probing was not promising

`PDDoc.GetJSObject()` returned an object, but:

- normal reflection/introspection was unstable
- `execMenuItem` was not exposed
- `beginPriv` / `endPriv` were not exposed
- `trustedFunction` was not exposed
- probing `preflight` produced COM access errors

Conclusion: JavaScript is not the most promising next step from Python COM alone.

## Exact Internal Identifiers Found

The desired fixup is real and was precisely identified.

### Workflow

- Title: `Optimize for Web and Mobile`
- Relative path: `Action04.sequ`
- Workflow ID: `CB7C61DD0538D7707B0D598AB312A0F`

### Preflight fixup/profile entries inside `Action04`

Two `CALS:Preflight` steps were found in the workflow:

1. `Convert to sRGB`
2. `Embed missing fonts`

For `Embed missing fonts` the following internal values were confirmed:

- Profile name: `Embed missing fonts`
- Dict key: `P_7_Embedmissingfonts`
- Fingerprint: `P9db551f478f00782f340fa57fd08cf08`

## Files And Registry Evidence

### Sequence files

- `C:\Program Files\Adobe\Acrobat DC\Acrobat\Sequences\ENU\Action04.sequ`
- `C:\Program Files\Adobe\Acrobat DC\Acrobat\Sequences\CHS\Action04.sequ`

Both include `CALS:Preflight` commands referencing `Embed missing fonts`.

### Preflight data files

- `C:\Program Files\Adobe\Acrobat DC\Acrobat\plug_ins\Preflight\Actions.kfp`
- `C:\Program Files\Adobe\Acrobat DC\Acrobat\plug_ins\Preflight\Acro2015.kfp`

These also contain the same `P_7_Embedmissingfonts` and fingerprint entries.

### Registered workflows

Registry root:

- `HKCU\Software\Adobe\Adobe Acrobat\DC\Workflow\cRegistered`

Relevant registered entry:

- key: `c3`
- title: `Optimize for Web and Mobile`
- relative path: `Action04.sequ`
- ID: `CB7C61DD0538D7707B0D598AB312A0F`

### `.sequ` association

Registry association:

- `HKCR\.sequ -> Acrobat.Sequence`
- `HKCR\Acrobat.Sequence\shell\Import_Action\command`

Command value:

```text
"C:\Program Files\Adobe\Acrobat DC\Acrobat\Acrobat.exe" "%1"
```

Interpretation: opening a `.sequ` file appears to import the action into Acrobat, not play it.

## Important Negative Findings

The probe tried all of the following as possible `MenuItem...` candidates and found them unavailable:

- `Optimize for Web and Mobile`
- `Action04.sequ`
- `Action04`
- `CB7C61DD0538D7707B0D598AB312A0F`
- `Convert to sRGB`
- `PPP_ConverttosRGB`
- `Pf0dc5626578ed8a915e83ccc5a2a5184`
- `Embed missing fonts`
- `P_7_Embedmissingfonts`
- `P9db551f478f00782f340fa57fd08cf08`

Conclusion: these are not directly executable through the discovered `AcroExch.App.MenuItemExecute(...)` surface.

## Practical Conclusion

The Acrobat-side capability clearly exists, but the current public COM surface that was discovered is not enough to execute `Embed missing fonts` directly.

At the time this work was paused, the most plausible next directions were:

1. investigate undocumented `Acrobat.exe` workflow playback arguments
2. investigate an import-plus-play workflow path for registered `.sequ` actions
3. build a UI automation PoC using COM only for document setup and window activation

## Archived Files

This archive includes:

- `README.md` - investigation summary and findings
- `acrobat_probe.py` - current Python COM probe used for exploration
- `test_acrobat_probe.py` - helper tests for the probe's parsing and probe logic

## Status

Feature development paused.

No integration into `controller.py`, `view.py`, or `pdf_processor.py` was performed for Acrobat-based font embedding.
