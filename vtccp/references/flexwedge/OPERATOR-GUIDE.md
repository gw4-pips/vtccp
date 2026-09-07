# FlexWedge 1.0 operator guide

**Document version:** 1.0.0  
**Revision date:** 2026-09-06

Unzip the approved package and run `install.ps1` as the signed-in user; it installs program files below `%LOCALAPPDATA%\Programs\VCCS\FlexWedge` and creates a desktop shortcut without administrator rights. Mutable configuration and sessions remain separately below `%LOCALAPPDATA%\VCCS\FlexWedge`, so upgrades do not overwrite them. Start FlexWedge, choose the ASR-P35U COM port shown in Device Manager, connect, then use **Single trigger**. The release deliberately accepts one reader and one tag for a trigger.

The application loads its sample configuration on first use and persists the selected port, profile fields, and optional GCP XML path at `%LOCALAPPDATA%\VCCS\FlexWedge\flexwedge.config.json`. Select a bundled GCP table by leaving the custom path blank, or browse to a valid customer-controlled GS1 GCP Prefix Format List XML and save configuration. A bad/missing table is displayed as an error; it does not manufacture a result and GCP remains `NotChecked`.

For training or regression, paste EPC hex into **Manual EPC input** and process it. Choose Raw EPC, GTIN + serial, Delimited mapped fields, or Complete decoded result. The fields/header box accepts ordered `Field=Header` entries separated by semicolons, so output columns are selected without worksheet coordinates. Configure prefix, suffix, delimiter, and append key, then save. Use Copy to Clipboard or first focus an external application, return to FlexWedge, and use Send to Previously Focused Application. FlexWedge continuously tracks the real external foreground window and confirms it regained focus before pasting.

Every processed read is also appended automatically to `%LOCALAPPDATA%\VCCS\FlexWedge\sessions\current.jsonl`. Use Export to save the current in-memory session as CSV or JSONL. In VeriWedge, paste one GS1 element string or Digital Link barcode, then validate the one current RFID EPC. This is a comparison aid, not verifier automation.

ASR acquisition requires `AsReaderP3xU.dll` supplied under the vendor's redistribution terms. A build made without it says hardware support is unavailable; it does not simulate a reader.

Running `uninstall.ps1` removes only program files and preserves configuration/session data. Use `uninstall.ps1 -PurgeUserData` only when you intentionally want to delete retained operator data.