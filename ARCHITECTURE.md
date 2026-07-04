# Architecture

The project is organized as a modular desktop application.

```text
app/                         Application entry points and version metadata
app_core/config/             Settings, paths, theme constants
app_core/services/           Registry, validation, generation, PDF/output services
app_core/engines/modules/    Finalized document generation engines
resources/templates/         Excel input templates
resources/assets/            Logo and branding assets
ui/                          Dashboard UI
packaging/macos/             macOS build/uninstall scripts
packaging/windows/           Windows build/installer scripts
tests/                       Smoke tests
```

The document engines are kept isolated from the UI so future documents can be added with minimal UI changes.


## Purchase Order Module

Purchase Order is implemented as a dedicated engine wrapper under `app_core/engines/modules/purchase_order/` and reuses the premium Invoice-style DOCX/PDF rendering engine with a Purchase Order profile. This preserves the proven header, table, summary and PDF generation approach while keeping PO-specific templates, registry metadata, output folder and UI card separate.
