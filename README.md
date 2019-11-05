# CF ToolBox

This application imports SciFiner's output (like RTF, SDF, etc) and output as ChemFinder database.

## Requirements

- Windows
- Microsoft Word
- ChemOffice/ChemDraw Professional 2019 because the following application is running behide this application.
	- ChemDraw
	- ChemFinder
	- ChemScript

## Install

- Install using `CFToolBoxSetup.msi` file.
- `CF ToolBox` is shown in Windows start menu.

## Gettting started

### How to create ChemFinder database from SciFinder outputs

1. Launch `CF ToolBox` application from start menu of Windows.
2. Menu > *Buid* > *Create from ...* > *SciFinder's RTF and SDF*.
3. Select files to import and click *OK*.
4. Specify the name of new ChemFider database.
5. New line is created and wait until showing finished.
6. Select the line and *File* > *Open* to launch ChemFinder or double-click the line to open created databse.
- You can select *CAS ONLINE*, *Compound name list*, *SMILES list*, and *ChemFinder* instead of  *SciFinder's RTF and SDF*.
- The *list* is a file that describes one compound name/SMILES per line.
- You can stop to import by selecting the checkbox for the task you want to stop and Menu > *Task* > *Kill*.

### How to append to existing database

1. Menu > *File* > *Add ...*.
2. Select ChemFinder database file (ie, `.cfx` file).
3. Select the newly created line.
4. Menu > *Build* > *Append ...*
5. Select files to append and click *OK*.
6. Wait until showing finished.

### How to manipulate database

1. Menu > *File* > *Add ...*.
2. Select ChemFinder database file (ie, `.cfx` file).
3. Select the newly created line.
4. Menu > *Build* > *Manipulate ...*.
5. Select `*Generate Structre from* *name/SMILES/InChI*`/`Generate SMILES from Structure`/`Clean up Structure`/`Scafford Structure`
6. Wait until showing finished.
- The structure used in `Scafford Structure` is specified in Menu > *File* > *Settings ...* > *Scaford Structure file*.
