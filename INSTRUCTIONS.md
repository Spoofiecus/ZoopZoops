# How to Install and Use the Layout Generator Macro in CorelDRAW

Here's how to install and use the `LayoutGenerator.gms` macro in your CorelDRAW application.

## Installation

1.  **Open the VBA Editor:**
    *   In CorelDRAW, go to the `Tools` menu.
    *   Select `Scripts` (or `Macros` in older versions).
    *   Click on `Script Editor` (or `Macro Editor`). This will open the VBA Editor window.

2.  **Import the Macro File:**
    *   In the VBA Editor, you'll see a "Projects" panel, usually on the left.
    *   Right-click on `GlobalMacros` in the project list.
    *   Select `Import File...`.
    *   Navigate to where you saved the `LayoutGenerator.gms` file and select it.
    *   The `LayoutGenerator` module will now appear under the `Modules` folder within `GlobalMacros`.

## How to Use

1.  **Prepare Your Document:**
    *   Open the CorelDRAW document you want to work in.
    *   Create or import the rectangular/square label design you want to lay out.
    *   **Select only that single label shape.**

2.  **Run the Macro:**
    *   Go to the `Tools` menu.
    *   Select `Scripts` (or `Macros`).
    *   Click on `Run Script` (or `Macro Manager`).
    *   In the `Macros in:` dropdown, select `GlobalMacros`.
    *   You should see `LayoutGenerator.CreateCutlineGrid` in the list of macros.
    *   Select it and click `Run`.

3.  **Enter Batch Height:**
    *   A dialog box will appear asking for the "batch height in mm".
    *   Enter your desired height (e.g., 300) and click `OK`.

4.  **Done!**
    *   The script will create a new layer named "Cutlines" containing the grid. The grid will be a single grouped object.

## Important Note on Cutline Color

*   The script is configured to use a spot color named **`CutContour`**.
*   For the script to work correctly, you **must** have a spot color with this exact name in your document's color palette. If your cutting software uses a different name (like "CutLine" or "PerfCutContour"), you can easily change it by editing the macro.
*   To edit the macro, open the VBA editor again, double-click the `LayoutGenerator` module, and change the `CUTLINE_COLOR_NAME` constant at the top of the script.
