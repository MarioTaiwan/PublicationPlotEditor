# Publication Plot Editor 

This program is designed to streamline the figure preparation process: Instead of having to edit every plot again, you just have to copy the plot from Excel and apply a Master style that handles colors, fonts, etc.

## Run

Extract this archive on Windows, put the extracted folder on the MATLAB path, and run:

```matlab
clear PublicationPlotEditor
PublicationPlotEditor
```

Designed for MATLAB R2021a+; actual MATLAB/CorelDRAW execution and EMF import cannot be tested in this build environment.

This is designed for Excel: Make a plot and use the Series information in the editor
<img width="1920" height="1080" alt="image" src="https://github.com/user-attachments/assets/bcfcd9dc-73ae-4cd2-9816-6237e8d9822c" />



## Setting up a reusable font master

1. In **Axes & ticks > X axis**, enter the required **Label font (pt)** and **Tick font (pt)**. Repeat in **Y axis**. The sizes are independent.
2. In **Panel**, enable **CorelDRAW EMF font correction** and select **Calibrate X/Y EMF font sizes...**. Four entries are available: X ticks, Y ticks, X label, and Y label.
3. The dialog shows the target font size from the corresponding axis tab. To calibrate, export an EMF with the current factors, open it in CorelDRAW, select representative labels, and enter their *observed sizes* in the matching fields. Click **Calculate from measurements**, then **Apply corrections**. Alternatively type the four multipliers directly.
4. Click **Save master** in the main toolbar. Your X/Y font sizes, four EMF multipliers, correction switch, panel size, note field, and other master appearance settings are stored in the JSON.
5. In a later session choose **Load master** to restore them, paste Excel SERIES formulas, and export your plots. Figure-specific recipes also preserve the same values.

For each text category, the correction formula is:

```
new multiplier = old multiplier * desired point size / measured CorelDRAW point size
```

Measurements must be from an EMF exported using the **current** factors and target sizes. Leave a measured-size field blank to keep that category's factor unchanged. If, for example, you requested a 7 pt Y-axis label and CorelDRAW reports 6.16 pt, the Y-label factor is multiplied by 7/6.16; tick corrections are unaffected.

The correction factors affect **EMF only**. PDF and editable MATLAB FIG files keep the requested MATLAB font sizes. Note that EMF/CorelDRAW font metrics are renderer- and font-dependent; a calibration is empirical, not universally valid.

## Backwards compatibility

- Load existing v0.6/v0.6.1 master and figure JSON files. The legacy shared tick/axis-label factors are copied into the four corresponding X/Y settings, retaining legacy custom values. The saved output uses schema version 7.
- Existing figure data, source Excel SERIES formulas, line styling, marker colors, fixed RGB palette choices, axes/box behavior, and panel geometry are unchanged.
- When you save a master, font sizes typed in either axis tab are captured even if the separate **Apply X/Y axis settings** button was not pressed. Other axis edits such as limits and tick sequences still require applying their tab settings.

## Key limitation

This is statically checked code, **not yet tested in MATLAB or CorelDRAW 19**. Check that saving/reloading a v0.7 master restores all four calibration values and X/Y label sizes, and verify import font sizes using one test EMF.
