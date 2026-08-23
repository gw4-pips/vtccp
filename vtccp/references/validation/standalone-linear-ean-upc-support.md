<!-- v1.2 | Revised 2026-08-22 -->

# Standalone EAN/UPC TruCheck Report Support

## Evidence boundary

This note describes the implemented **parser and report behavior**, not a claim
that a live DMST or Webscan TruCheck capture has been obtained. No capture is
fabricated for either verifier. The automated examples are intentionally
labelled parser-shape fixtures and are not retained as source evidence.

The standalone report path accepts one literal verifier symbol row and one
literal ISO/IEC 15416 Verification Grades row. EAN/UPC is handled exactly like
any other GS1 symbology: the normal two-parser behavior remains in force, and
a literal verifier Data Format Check can contain a GTIN field that is not 14
digits and a separately reported check digit. Those literal rows are preserved
unchanged. The report never locally recomputes a check digit. A Data Format
Check appears only when the correlated verifier HTML supplies its own table;
otherwise the report states that it is unavailable.

A report is classified as standalone only when its explicit verifier summary
names EAN-13, UPC-A, EAN-8, or UPC-E, exactly one overall-grade display is
present, and the HTML contains no matrix, QR, ECC, or other 2D characteristics.
This classification describes the verifier evidence shape only; it does not
change the GS1 parser policy or suppress the normal two-parser report layout.

## Variant classification

| Variant | Status | Current report behavior | Evidence / decision |
|---|---|---|---|
| EAN-13 | Supported parser/report shape | One literal EAN-13 summary row, one literal ISO/IEC 15416 grade row, native GS1 DFC rows when present, and the normal two-parser behavior. | A non-14-digit GTIN and its separate check digit are retained exactly as reported. Needs a controlled native capture before it can be called live-validated. |
| UPC-A | Supported parser/report shape | One literal UPC-A summary row, one literal ISO/IEC 15416 grade row, native GS1 DFC rows when present, and the normal two-parser behavior. | A non-14-digit GTIN and its separate check digit are retained exactly as reported. Needs a controlled native capture before it can be called live-validated. |
| EAN-8 | Supported parser/report shape | One literal EAN-8 summary row, one literal ISO/IEC 15416 grade row, native GS1 DFC rows when present, and the normal two-parser behavior. | The verifier, rather than VCCS, supplies its GTIN/check-digit interpretation. Needs a controlled native capture before it can be called live-validated. |
| UPC-E | Supported parser/report shape | One literal UPC-E summary row, one literal ISO/IEC 15416 grade row, native GS1 DFC rows when present, and the normal two-parser behavior. | No local expansion or check-digit policy is applied. Needs a controlled native capture before it can be called live-validated. |
| UPC/EAN add-ons | Separate product decision | The base-symbol standalone classifier does not claim add-on support. | Capture the verifier's symbology label, decoded-data representation, and grade-row shape first. |
| Missing verifier fields | Unavailable | Missing cells remain absent in the model and show as unavailable in the report. | Never infer an omitted field from the symbology or grade. |
| Unexpected decoded length | Unavailable | The literal decoded data is retained, but no local validity or check-digit result is added. | Requires a native verifier result or a separate approved parsing policy. |
| Missing verifier DFC table | Unavailable | The report displays `DATA FORMAT CHECK UNAVAILABLE — NOT PRESENT IN TRUCHECK HTML`. | No local check-digit recomputation is permitted. |

## Required controlled capture before live validation

On the Windows USB verifier workstation, retain each original HTML report and
its sibling barcode image without editing either. Capture EAN-13 and UPC-A
first, then EAN-8 and UPC-E if available. For each scan, compare the literal
symbology, decoded data, Verification Grades cells, Data Format Check presence,
and embedded-image behavior with the rendered VCCS PDF.
