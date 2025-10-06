# invoice-pegaz-flow — Project Brief

## Metadata

• Title: invoice-pegaz-flow — Project Brief
• Version: 0.1.0
• Date: 2025-10-05
• Status: Draft
• Project start: Jun 15, 2025
• Development status: actively developed

## Authorship and Ownership

• Author: Mariusz Grzelak
• Contact: mariusz.grzelak84@gmail.com
• Project owner: Mariusz Grzelak (individual project, not corporate)
• © 2025 Mariusz Grzelak. All rights reserved.

## Project Goal (Summary)

Creation of a lightweight system for warehouse management and purchase/sales documents with emphasis on process automation, invoice import (PDF), local data storage (offline-first) and integration with Google Drive and Excel export.

## Scope

• Goals: warehouse inventory management, transaction recording (purchases/sales), PDF import, report export, inventory alerts, backup and document printing.
• Non-goals (for now): full accounting/ERP, extensive multi-company functionality, multi-role authorization.

## Key Features

• Warehouse: inventory levels (StockLevelsTable), movements (StockMovements), alerts (StockAlertsCard), quick search and filters (SearchAndFilter, SmartSearchCommand).
• Products/categories/suppliers: ProductManager, CategoryManager, SupplierManager.
• Operations: initial inventory, purchases, sales, printing (PrintableForms), user preferences.
• Import/export: PDF invoice import (PDFUploader, EnhancedPDFUploader, services/pdfParser.ts, services/advancedPdfParser.ts), Excel export (services/excel.ts).
• Integrations: Google Drive (GoogleDriveManager, services/googleDrive.ts), backup/synchronization (BackupAndSync, DataBackup).
• UX/Infrastructure: keyboard shortcuts, notifications (toast/sonner), UI protections (ErrorBoundary), lazy loading, query cache (React Query).

## Architecture (High Level)

• Stack: Vite, React, TypeScript, shadcn-ui, Tailwind CSS, React Router, React Query.
• State and data: Zustand (stores/inventoryStore.ts) + IndexedDB/Dexie (services/database.ts) with transactions and event-sourcing of warehouse movements.
• Layers: components/ (UI), services/ (I/O: PDF, Excel, Drive, DB), hooks/ (screen logic), types/ (contracts), lib/utils.ts (helpers).
• External integrations: Google Drive (googleapis), Excel (xlsx), OCR (tesseract.js – if used), PDF parsing.

## Data and Persistence

• Local IndexedDB database (Dexie) – offline-first mode, transactions for complex operations (purchases/sales), default data initialization.
• Export/backup to files and integration with Google Drive (versioning on Drive side).

## Security and Backups

• User data stored locally (browser).
• Security backups and restoration via Google Drive and local files.
• Import validations (dates, duplicates, item consistency).

## Automation (Priorities)

1. PDF invoice import → automatic item parsing → entry into purchases/sales → inventory updates (validation, duplicates, VAT ID, dates).
2. Export of warehouse and financial reports to Excel (inventory, margins, turnover, rotation).
3. Minimum inventory alerts and order suggestions (thresholds, lead time).
4. Backup/synchronization to Google Drive (versioning, restoration).
5. Document printing (delivery notes/receipts/simplified invoices) and labels with codes.
6. Code scanner and quick receipt/issue operations.

## Success Metrics (Examples)

• Reduction of manual data entry time by ≥50% through PDF import.
• Reduction of warehouse shortages (alerts) and improvement of inventory rotation.
• Reliable environment restoration from backups (backup/restore tests).

## Assumptions and Limitations

• Single-location and local data (for now).
• No full ERP; accounting integrations outside MVP scope.
• Private repository, individual work.

## Roadmap (Outline)

• MVP: PDF import, purchase/sales recording, inventory, Excel reports, alerts, backup.
• Next steps: document and label printing, code scanner, extended analytics.
• Research: external integrations (supplier APIs), permissions/roles.

## Risks and Mitigations

• PDF quality (parsing risk) → fallback to semi-automation/editing.
• Local data consistency → transactions, validations, integrity audits.
• Local data loss → regular backup (Drive + local).

## License / IP

• Unless otherwise specified: © 2025 Mariusz Grzelak. All rights reserved.
• In case of adding a license in the repository, the license record takes precedence.

## Change Log

• 0.1.1 (2025-10-06): Corrected project start date to Jun 15, 2025.
• 0.1.0 (2025-10-05): First version of professional project brief.
