CATIA Advanced VBA Automation Suite

A professional-grade automation framework for CATIA V5, built using VBA, designed to streamline repetitive CAD tasks across part creation, assembly generation, drawing creation, and full project packaging.
This suite is ideal for design engineers, automation developers, and organizations aiming to reduce manual modeling time and enforce consistent CAD standards.

🚀 Features
1. Parametric Part Generator

Creates CATParts using Excel-driven parameters

Supports sketches, pads, pockets, fillets, holes, patterns

Enforces naming standards and clean tree structure

2. Automatic Assembly Builder

Generates CATProducts using a set of part inputs

Applies constraints: coincidence, offset, contact, angle

Builds multi-level assemblies programmatically

3. Engineering Drawing Creator

Creates CATDrawings based on templates

Auto-places front, top, side, and isometric views

Adds dimensions, centerlines, and title block data

Supports batch drawing creation

4. Project Export & Packaging

Exports complete CAD package:

Parts

Assembly

Drawings

Auto-generates delivery folders

Supports IGES, STEP, PDF, and CATIA native formats

5. Utility Tools

Screenshot capture tool

Logging system

Error handling & message reporting

Execution time tracker

📂 Project Structure
CATIA-Advanced-VBA-Automation/
│
├── vba/
│   ├── GenerateParts.bas
│   ├── BuildAssembly.bas
│   ├── CreateDrawings.bas
│   ├── ExportPackage.bas
│   └── Utilities.bas
│
├── excel/
│   └── parameters.xlsx
│
├── templates/
│   ├── PartTemplate.CATPart
│   ├── DrawingTemplate.CATDrawing
│   └── AssemblyTemplate.CATProduct
│
└── output/
    ├── parts/
    ├── assembly/
    ├── drawing/
    ├── package/
    └── screenshots/

🛠️ Prerequisites

CATIA V5 R19 or later

Microsoft Excel

VBA enabled in CATIA (Tools → Macros → Security)

▶️ How to Use
1. Load Macros

Open CATIA

Go to Tools → Macros → Macros…

Add the vba folder to your library

2. Configure Excel Parameters

Modify excel/parameters.xlsx to set sizes, features, and assembly definitions.

3. Run Automation Modules

GenerateParts.bas → Creates CATParts

BuildAssembly.bas → Builds CATProduct

CreateDrawings.bas → Generates CATDrawings

ExportPackage.bas → Outputs delivery package

📦 Download

The full project is available as a ZIP here:
👉 CATIA-Advanced-VBA-Automation.zip

🤝 Contributing

Feel free to fork this repository and submit pull requests.
Suggestions for new automation modules are welcome.

📜 License

This project is released under the MIT License.