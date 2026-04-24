# AutoCAD Board Sorter

AutoCAD .NET add-in for sorting cabinet/furniture board solids by layer and thickness.

Command:

```text
BDSORT
```

What it does:

- scans selected `3DSOLID` objects, or all model-space solids if you press Enter;
- calculates sorted dimensions: length, width, thickness;
- groups parts by `Layer + Thickness`;
- prints a report in the AutoCAD command line;
- writes a CSV next to the DWG: `<drawing>_boards.csv`.

## Build

This project targets `.NET 8` because AutoCAD 2025 loads .NET 8 plugins.

Set the `ACAD` environment variable to the AutoCAD install folder containing `AcMgd.dll`, `AcDbMgd.dll`, `AcCoreMgd.dll`, and `AcDbMgdBrep.dll`.

Example:

```bat
set ACAD=C:\Program Files\Autodesk\AutoCAD 2025
build-release.bat
```

Or open `AutoCAD_BoardSorter.sln` in Visual Studio and build Release.

## Load In AutoCAD

1. Run `NETLOAD`.
2. Pick `src\AutoCAD_BoardSorter\bin\Release\AutoCAD_BoardSorter.dll`.
3. Run `BDSORT`.

## Autoload

Quick registry way:

1. Build the project.
2. Run `NETLOAD` once and load `AutoCAD_BoardSorter.dll`.
3. Run `BDREGISTER`.
4. Restart AutoCAD.

To remove registry autoload, run:

```text
BDUNREGISTER
```

Bundle way:

```bat
set ACAD=C:\Program Files\Autodesk\AutoCAD 2025
install-bundle.bat
```

This installs `AutoCAD_BoardSorter.bundle` into one or both available plugin folders:

```text
%APPDATA%\Autodesk\ApplicationPlugins
%PROGRAMDATA%\Autodesk\ApplicationPlugins
```

## Notes

The main algorithm is adapted from the SolidWorks macro:

- choose the largest planar face;
- build a local coordinate basis from that face;
- collect candidate angles from face edges;
- rotate axes and minimize the projected XY area;
- take the three final dimensions and sort them descending.

For AutoCAD, the add-in projects BRep vertices onto candidate axes. If a solid has no usable planar BRep data, it falls back to `GeometricExtents`.
