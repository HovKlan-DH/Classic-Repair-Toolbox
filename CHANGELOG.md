# 2.5.0, 2026-September-?
-------------------------
- **Application**
  - Added "Workbooks" for tracking repairs
    - Added `--workbooks-root` command-line argument to override default folder location
    - Added ability to capture and attach oscilloscope images to worklog (via component info popup)
  - Added ability to add new component in "Contribute" tab
  - Added ability to delete an existing component in "Contribute" tab
  - Added application crash logging to prevent losing info when overwriting logfile
  - Changed contribution requirements via "Contribute" tab to version `2.5.0` or newer (server backend requires it)
  - Changed info banner to specifically show if BETA is used as source during data updates
  - Fixed manifest file performance via server-side compression, resulting in faster online data checks
  - Fixed timeout issue when fetching manifest file
  - Fixed test case configuration to exclude them when compiling using `dotnet publish`
  - Refactored "Contribute" tab and server backend review process
  - Refactored parts of project for better maintainability and added more test cases (2827 cases and ~54% coverage total)
  - Refactored documentation


# 2.4.0, 2026-August-27
-----------------------
- **Application**
  - Added "External tooling availability" section in "Configuration" tab
  - Added BOM (Bill of Material) list to be printed in "Overview" tab
  - Added selector to include/exclude component from being printed in the two lists in "Overview" tab
  - Added link to YouTube videos in "About" tab
  - Added help icon in various places, linking to GitHub Wiki page
  - Added KiCad data initializing indicator in "Schematics" tab
  - Added commandline option for `--simulate-update`
  - Changed MiniPro IC test from experimental to an official feature
  - Changed MiniPro IC test is embedded in component info popup and has several UI changes
  - Changed mouse location is keeping logical position when zooming
  - Changed component info popup window opens in same position, if only a single window gets opened
  - Changed saving the user preference file now uses atomic-swap instead of direct-write
  - Changed build strategy, so a `Debug` build no longer has special application behaviour
  - Fixed KiCad trace overlay was slow to zoom and to hover on boards with many components
  - Fixed KiCad rectangular and oval pads ignored their rotation
  - Fixed KiCad pads on bottom side components were placed mirrored
  - Fixed KiCad traces would not highlight when selecting a component on the bottom side of a board
  - Fixed KiCad hierarchical sheets were loaded twice, listing the same schematic under two different "CAD name" values in the log
  - Fixed missing page overflow scrollbar in "Configuration" tab
  - Fixed missing page overflow scrollbar in "Contribute" tab
  - Fixed closing MiniPro IC test would require two clicks on test IC button to go to test again
  - Fixed MiniPro IC test could show the result from the previously tested IC, if switching to another component while a test was still finishing
  - Fixed MiniPro IC test could freeze the application for a long time when a test produced a large amount of output
  - Fixed oscilloscope `T/DIV` values in the nanosecond range could be matched to a wrong supported value
  - Fixed oscilloscope image dump could leave the connection out of sync, so all following commands would read the previous commands response
  - Fixed some buttons would reveal text below when hovering with mouse
  - Fixed Linux proper maximizes the application window at launch, if this was its last state
  - Fixed delete routine of orphan/unused files could in rare cases wipe entire data-root folder (CRT data only at risk)
  - Fixed opening a file from application will not blindly _execute_ it and leave it for system to find appropriate means, but only _open_ whitelisted extensions
  - Fixed an intermittent failure where a board could load empty at startup when data validation was enabled
  - Refactored several parts of project for more efficient AI development, including unit/UI test coverage (1074 cases and ~31% coverage total)


# 2.3.0, 2026-July-24
---------------------
- **Application**
  - **New**
    - Added __experimental__ possibility to test IC's with a MiniPro programmer (TL866/T48/T56)
      - "Configuration" tab has a couple of options for enabling this (not enabled per default)
      - Read [MiniPro programmer documentation](https://github.com/HovKlan-DH/Classic-Repair-Toolbox/wiki/MiniPro-programmer-how%E2%80%90to) to see how to use this
    - Added show of proper BETA version number, if applicable
  - **Changed**
    - Increased "Feedback" upload timeout from 1 minute inactivty to 5 minutes
  - **Fixed**
    - Mouse wheel zoom now scales with reported scroll delta instead of fixed step per event, fixing overly coarse/aggressive zooming on some platforms
- **Data**
  - **New**
    - Added **Commodore Plus/4**
    - Added **Commodore 16**


# 2.2.0, 2026-May-21
--------------------
- **Application**
  - **New**
    - Added option to download data from BETA source
      - Should only be used in agreement with developer, as data can be incomplete or even broken for some parts
      - Used to test new functionalities, and should not be used as a normal data source!
  - **Changed**
    - Banner will show filenames when updating data from online source
    - "Clear All" button will not be shown in collapsed "Important signals" when no selected signals
    - "Clear All" button will not be shown in "Netlist names" when no selected signals
  - **Fixed**
    - Clean-up of orphan and non-used files did in some cases not delete all files in scope
    - KiCad calibration points would be reset after using component label editor
- **Data**
  - **New**
    - Added **ZX Spectrum Issue 4B**


# 2.1.0, 2026-May-9
-------------------
- Fixed an exception error in Linux, preventing application from running
- Added option for `Delete orphan and non-used files` in "Configuration" tab


# 2.0.1, 2026-May-7
-------------------
- Fixed online data update could incorrectly use an older local main Excel data file instead of the newer one from online source


# 2.0.0, 2026-May-7
-------------------
- Fixed component labels was shown on top of panels
- Fixed selected component did not blink in thumbnails
- Fixed it should not zoom schematic image when inside a panel
- Fixed "Netlist names" panel should not show in schematic images without KiCad data
- Fixed thumbnail opacity can be 100% (pure solid)
- Added `Important signals` panel in "Schematics" tab
- Added visualization of KiCad traces on opposite side
- Added visualization of KiCad zones
- Added visualization of KiCad traces also in schematic images (not only PCB images)
- Changed update banner notification when main Excel data file has newer version
- Changed "Clear All" button is now always visible in "Netlist name" panels, if it has selected traces
- Changed it is possible to pan, so panels in "Schematics" tab does not hide the schematic below it
- Changed KiCad calibration mapping to use interactive trace-map instead of manual calibration points
- Breaking changes:
  - Changed data format for KiCad files
    - Now using original raw KiCad data files instead of a JSON file with converted data
  - Changed data format for highlighted component bounds
    - Moved from board Excel sheet to JSON file
  - Changed data format for KiCad calibration points
    - Moved from board Excel sheet to JSON file

> [!CAUTION]
> Due to breaking data format changes, the major version has been increased to reflect this is important.
> No more data updates will be given for versions **below 2.0.0**, and as future versions may break current
> data, then you are recommended to update.


# 1.8.0, 2026-April-25
----------------------
- Fixed some Linux distros could not open folder location of data/log/settings #98
- Added icon for showing "fetch data from online source" status and to refresh instantly


# 1.7.0, 2026-April-18
----------------------
- **Application**
  - Fixed zoom-in keeps its mouse/image position at first zoom-in
  - Fixed "Clear" button did not clear component filter input field
  - Fixed board labels (if visible) are shown on top of traces and pads
  - Fixed (again) `ESCAPE` did not close active component info popup in Linux, if set to "Open multiple windows for popup" #84
  - Fixed horizontal splitter in "Schematics" tab is shown in correct position at application launch
  - Fixed (experimental) that **macOS** trackpad was not working properly
  - Added "Global settings" panel in "Schematics" tab and moved some functionalities to it
  - Added option how KiCad traces should be highlighted in new "Global settings" panel
  - Added option to highlight first pin on component, if schematic has KiCad data
  - Added user configurable theme colors (editable through configuration file)
  - Changed all components can be selected or highlighted in "Schematics" image, not being dependent on selected categories any more
  - Changed KiCad calibration points only includes pads and reference texts (only relevant for contributors using KiCad)
  - Changes in component label editor:
    - Added component multi-select for move and resize
    - Added component snap-align on move and resize
    - Added component duplication with the letter `D`
    - Added an "undo" and "redo" stack with `CTRL`+`Z` and `CTRL`+`Y`
    - Added banner to show when in this mode
    - Changed component search filter will now also search for components when inside component label editor
    - Changed it will disable online data updates if any changes have been applied, to prevent data loss of local file
    - Changed it will reuse category when creating new component
    - Changed it is no longer possible to change to another schematic, board or hardware
    - Changed YouTube video for showcasing the component label editor:
      - [How to use component label editor in CRT](https://youtu.be/u-UkD-m4Z6o)
  - Removed "Data contribution" checkboxes in "Configuration" tab
    - Same functionality now applies when checking the "Enable contributor mode" in the "Global settings" panel
  - Refactored parts of code for better performance and minor text changes
- **Data**
  - **Commodore 64**
    - **KU-14194HB**
      - Added new board (has KiCad data)
    - **250407**
      - Reexported images from KiCad in a better quality
      - Relabeled all images
    - **250469**
      - Reexported images from KiCad in a better quality
      - Relabeled all images

> [!CAUTION]
> Due to major update of Avalonia UI then Windows 32-bit is no longer supported 😥
> The newest available .NET LTS (Long-Term Support) will still be included for all 64-bit packages.


# 1.6.0, 2026-April-10
----------------------
- Added interactive KiCad traces and netlists for these boards:
  - **Commodore 64**
    - **250407 (long board)**
    - **250469 (short board)**
- Added YouTube video for showcasing interactive traces functionalities:
  - [How to use interactive traces in CRT](https://youtu.be/Y55nC_gJbH4)
- Added "Board settings" panel in "Schematics" tab
- Added "Print component list" in "Overview" tab
- Changed it will only show selected categories/components in "Overview" tab
- Changed it now highlights the component when hovering mouse over it
- Refactored parts of code for better performance


# 1.5.0, 2026-April-5
---------------------
- Fixed it will not show "oscilloscope disconnected" at application launch, if not set to auto-connect
- Added label editor in "Schematics" tab
- Added YouTube video for showcasing label editor functionalities:
  - [How to use component label editor in CRT](https://youtu.be/jmYrz2VB3ig)
- Added more checks for the data validator
- Changed "Contribute" tab:
  - Added UUID per row in Excel data files
  - Added "File location" drop-down
  - Changed "File" does no longer include relative path
- Refactored and mitigated some potential security risks


# 1.4.0, 2026-March-31
----------------------
- Fixed changing to another hardware will clear the filter component search
- Fixed horizontal misalignment of rows in component contribution window
- Added "Oscilloscope" tab
- Component info popup:
  - Added oscilloscope sync when viewing oscilloscope images
  - Added oscilloscope labels in the top-right corner for the large image
- Added YouTube video for showcasing oscilloscope functionalities:
  - [How to sync oscilloscope in CRT](https://youtu.be/CbTh1FFp3tU)
- Changed minor UI and text changes


# 1.3.0, 2026-March-20
----------------------
- Fixed deselecting component from a filtered list did not remove highligting in schematic image
- Fixed vertical table header/cell misalignment in "Overview" tab
- Fixed thumbnails in component info popup also shows a lighter background color on transparent images
- Added "Contribute" tab
- Added (experimental) a couple of YouTube videos for showcasing specific functionalities:
  - [Short introduction](https://youtu.be/B36etv0zOMo)
  - [How to do manual traces](https://youtu.be/JUNXeCHsrME)
- Changed icon format in Linux from `.ico` to `.png`
- Changed display name in **Linux** and **macOS** from `Classic-Repair-Toolbox` to `Classic Repair Toolbox`
- Changed some colors
- Refactored various minor code, UI and texts


# 1.2.0, 2026-March-14
----------------------
- Fixed it did show some board labels in wrong position in schematic image
- Fixed it did not save main region selection
- Fixed `ESCAPE` did not close active component info popup, if set to "Open multiple windows for popup" #84
- Added possibility to reorder schematic thumbnails with mouse drag/drop
- Added fullscreen mode - can also be enabled with `F11`
- Added image counter per region in component info popup
- Changed "Blink selected" checkbox is saved
- Changed region buttons and labels will get removed if no region data present
- Changed "About" tab
- Changed a few colors - still an ongoing thing
- Changed format for some logged events
- Changed application icon
- Changed display name in Windows from `Classic-Repair-Toolbox` to `Classic Repair Toolbox`
- Refactored various code and texts


# 1.1.0, 2026-March-9
---------------------
- Fixed traces are displayed correctly after a window resize #69
- Fixed transparent images did not show correct background color in component info popup #79
- Fixed `ESCAPE` and mouseclick outside of single-opened component info popup does not close window #68
- Fixed wrong reporting of version when submitting feedback
- Added traces will get focus when hovering it with mouse
- Added drop-down selector in component info popup deciding functionality for mouse scrollwheel
- Added option to undo deletion of traces
- Added `BUILDING.md` file for help on local build compilation
- Changed loading of main Excel data file to support strict binding between application and data versions #80
- Changed it will show component info popup on same monitor as main application
- Changed it will show mouse movement icon when panning image in "Schematics" tab
- Changed standard colors for traces floating panel
- Changed some themed colors (I believe this will be a long ongoing optimization)
- Changed format for some logged events
- Removed changelog from "About" tab


# 1.0.0, 2026-March-7
---------------------
- First public release


# 0.1.14, 2026-March-7
----------------------
- Changed format and text for log events
- Various refactoring in code base


# 0.1.13, 2026-March-7
----------------------
- Various refactoring in code base


# 0.1.12, 2026-March-6
----------------------
- Fixed that full width is used in "Schematics" tab
- Fixed manual trace could go beyond image area
- Fixed transparent images in component info popup where black text could not be seen
- Fixed CPU was not providing CPU architechture in the online check-in
- Added "Feedback" tab
- Added filtering of components in "Overview" tab
- Removed "Type" column from table "Overview" tab
- Removed "Help" tab (for now)


# 0.1.11, 2026-March-6
----------------------
- Fixed some compiler warnings


# 0.1.9, 2026-March-6
---------------------
- Fixed holding down right mouse button over component will not toggle selction in "Schematics" tab
- Added option for manual traces in "Schematics" tab
- Added component labelling in "Schematics" tab
- Added option to use development relases


# 0.1.8, 2026-March-5
---------------------
- Added content in "Resources" tab
- Added content in "Overview" tab
- Add component input filter
- Added region label in component info popup
- Changed region buttons are static in size and move along with splitter in component info popup
- Removed "Feedback" tab (for now)
- Various refactoring in code base


# 0.1.7, 2026-March-4
---------------------
- Added close button in component info popup
- Added border to splashscreen
- Added credits in "About" tab
- Added top-left labels in "Schematics" tab
- Added color for the two region buttons
- Added link to release notes on new version
- Changed region checkbox to buttons in component info popup
- Changed window title for component info popup
- Added that ENTER works similarly as SPACE in component info popup
- Various minor UI and text changes
- Various refactoring in code base


# 0.1.6, 2026-March-3
---------------------
- Added all data and (most) functionality to component info popup
- Added it will log CPU architechture in the online check-in
- Changed format and text for log events
- Removed checkbox for "Open popup in maximized mode" in "Configuration" tab


# 0.1.5, 2026-March-2
---------------------
- Changed from .NET10 to .NET6 (or newer)
- Changed from x64 only to both x86+x64 (Windows)


# 0.1.4, 2026-March-2
---------------------
- Fixed splashscreen is shown as early as possible in launch sequence
- Fixed it did not remember last used monitor
- Fixed some unhandled exceptions and added an "throw exception" as debug to logfile
- Fixed it auto-closes open component info popup(s) when closing application
- Added "Data" folder as part of GitHub project
- Added information for missing images that needs to be fetched via online sync
- Added refresh button after online sync and there were changes
- Added checkbox for data validation in "Configuration" tab (only relevant when contributing with Excel files)
- Added checkbox for debug log in "Configuration" tab
- Added closure of component info popup when clicking mouse outside of window
- Changed format and text for log events


# 0.1.3, 2026-March-1
---------------------
- Fixed highlighting in schematic image was bound to "PAL" region only
- Added functionality for blink button
- Changed right-click in schematic will toggle component selection
- Changed schematic thumbnails will dimm thumbnails without selected components


# 0.1.2, 2026-March-1
---------------------
- Renamed project to "Classic Repair Toolbox"


# 0.1.1, 2026-March-1
---------------------
- Fixed splashscreen would show a static text for version
- Added two checkboxes for component information popup in "Configuration" tab
- Added hovering on components in "Schematics" tab:
  - Label will show component name
  - Click to select component and open component information popup
  - Right-click to deselect component
- Changed appearance switch to be a checkbox instead in "Configuration"
- Various minor UI and text changes


# 0.1.0, 2026-February-28
-------------------------
- Initial pre-release
