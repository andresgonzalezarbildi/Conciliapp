# ConciliApp

A browser-based application for comparing transactions from an accounting system with cash or bank records. It can import real-world spreadsheets, review proposed matches, and export the results to Excel without sending accounting data to a server.

**Link to project:** <https://andresgonzalez.netlify.app/conciliapp/>

![Accounting Reconciliation interface](./conciliacion-contable.png)

## How It's Made:

**Tech used:** HTML5, CSS3, JavaScript, SheetJS, Lucide Icons, Web Workers

The application is built entirely on the client side. SheetJS is used to read XLSX, XLS, and CSV files and to generate the final Excel workbook, including summary, reconciled, possible, pending, original-data, and import-error sheets.

Accounting spreadsheets do not always begin on the first row or contain a single clean table. The importer detects possible headers, lets the user select a worksheet and table range, and provides manual mapping for date, description, amount, debit, and credit fields. Original values are preserved for export, while dates, amounts, and descriptions are normalized internally for comparison.

The reconciliation engine searches for one-to-one, one-to-many, and many-to-one matches. Each proposal receives a confidence score based on amount, date proximity, description similarity, and shared numeric references. Clear matches can be confirmed automatically, while ambiguous cases remain available for manual review and editing.

All processing takes place in the browser. Progress can be preserved locally, and an exported workbook can be opened on another computer to continue the reconciliation.

## Optimizations

The first version could potentially compare every transaction in one table with every transaction in the other. The engine was refactored to build indexes by date, sign, and amount, so it only evaluates transactions within the configured search window. In a synthetic test with 10,000 transactions per table, this reduced 100 million theoretical comparisons to approximately 820,000 candidate pairs.

The main processing task runs in a Web Worker when supported by the browser, preventing the interface from becoming unresponsive. The application also includes a workload estimate, separate limits for pair and grouping searches, cancelable processing, pagination, and warnings for potentially expensive configurations.

Additional passes operate only on unresolved transactions. Users can adjust tolerances for a more flexible search, exclude pending transactions outside a selected date range without deleting them, and restore those exclusions later.

## Lessons Learned:

The most difficult part was not reading an Excel file, but accepting spreadsheets with titles, auxiliary rows, multiple tables, and different sign conventions without relying on a rigid template. This made it important to combine automatic detection with visible manual controls.

I also learned that an accounting reconciliation should not treat an approximate match as a certainty. Separating confirmed, possible, and pending results — and explaining the score while allowing groups to be edited or removed — was just as important as the matching algorithm itself.

Finally, working with thousands of transactions showed the difference between optimizing calculation time and improving the complete user experience. Estimating workload, keeping the interface responsive, supporting cancellation, preserving progress, and limiting combinations are all parts of the same problem.

## Account workspaces and safer matching

The reconciliation engine now resolves repeated one-to-one amounts as a global assignment instead of accepting the first compatible pair. This gives description, numeric references, and date coherence enough weight to avoid swapped matches when several transactions share the same amount.

Mixed-sign groups and internal debit/credit offsets are disabled by default. If enabled, mixed-sign groups stay under manual review, and an internal offset can only be auto-confirmed when there is strong reversal evidence such as a shared reference or wording related to a return, reversal, or cancellation.

From the review screen, **Mover entre cuentas** can load another exported reconciliation or a locally saved account. Pending movements can be selected or dragged between accounts. The transfer preserves the original row and movement ID as audit metadata, stores both account states locally, and adds a **Movimientos transferidos** sheet to later exports.

Run the regression tests with:

```bash
npm test
```
## Traslado entre cuentas

La ventana **Mover entre cuentas** permite cargar otra conciliación y trasladar movimientos pendientes mediante selección o arrastrar y soltar. Antes de mover cada fila puede elegirse cómo debe quedar en la cuenta receptora: **Mantener**, **Débito** o **Crédito**. El cambio conserva el tipo original y el tipo aplicado en la hoja **Movimientos transferidos**.

