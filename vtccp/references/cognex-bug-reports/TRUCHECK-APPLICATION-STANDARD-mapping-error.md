# Cognex Bug Report — TRUCHECK.APPLICATION-STANDARD Integer Mapping Error

**Date:** 2026-08-05
**Reported by:** Product Identification and Processing Systems, Inc. (PIPS)
**Contact:** GW4

---

## Summary

The DMCC Reference documents the integer values for `TRUCHECK.APPLICATION-STANDARD` with 4 and 5 **swapped** relative to the actual firmware behaviour on the DM475V. The reference says 4=Auto and 5=Custom; the device returns the opposite.

---

## Reference document

| Field | Value |
|---|---|
| Document | DataMan Control Commands (DMCC) Reference |
| Revision | 26.1.0.27 |
| Date | 2026-04-21 |
| Parameter page | TRUCHECK.APPLICATION-STANDARD |
| Platforms listed | DM475V, DM8072V |
| Version (minimum) | 6.1.10 |

**Reference claims:**

| Integer | Label |
|---|---|
| 4 | Auto |
| 5 | Custom |

---

## Test device

| Field | Value |
|---|---|
| Model | DM475V-DPM |
| Serial / MAC suffix | 866D76 |
| IP | 10.10.10.4 |
| Firmware | 6.1.16_tc9 |
| State at test | Factory-reset, post-TruCheck calibration |

---

## Test procedure

1. Connected to the device in VCCS (TruCheck Verification Settings → Application Settings).
2. Observed the **Application Standard** dropdown showing **Auto**.
3. Queried via raw DMCC on port 23: `GET TRUCHECK.APPLICATION-STANDARD`
   → Device returned **`5`**
4. Changed the dropdown in the VCCS UI to **Custom** and saved.
5. Re-queried via raw DMCC: `GET TRUCHECK.APPLICATION-STANDARD`
   → Device returned **`4`**

---

## Actual behaviour

| UI label (VCCS) | DMCC GET returns |
|---|---|
| Auto | **5** |
| Custom | **4** |

---

## Expected behaviour (per reference)

| UI label | DMCC GET should return |
|---|---|
| Auto | 4 |
| Custom | 5 |

---

## Impact

Any third-party software reading or writing `TRUCHECK.APPLICATION-STANDARD` via DMCC and relying on the published reference will interpret and set the wrong mode. Code that sends `SET TRUCHECK.APPLICATION-STANDARD 4` intending to set **Auto** will actually set **Custom**, and vice versa.

---

## Request

Please confirm whether:

1. The DMCC reference revision 26.1.0.27 contains a documentation error (values for 4 and 5 are printed in the wrong order), **or**
2. The firmware has a known defect where the GET/SET values are inverted relative to the documented mapping.

In either case, please issue a corrected reference or firmware advisory so third-party integrators can rely on the published values.
