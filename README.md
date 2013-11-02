# Python automatisation wrapper around the NanoWrite software by nanoscribe GmbH
## Introduction
The NanoWrite software by nanoscribe GmbH is the user interface to their awesome 3D printer.
On its own, this software provides an intuitive user interface and comes with its own simple scripting language,
called GWL.

For most users, this will probably be enough already. But there are times where one needs to automate specific
processes like finding a specific marker, doing optical auto-focus or load a GWL file which was generated on the fly.

## Technical implementation
This wrapper automates the nanowrite software by simulating series of mouse and keyboard presses. Great care has
been taken to make the process as stable as possible. Matters are further complicated by NanoWrite being a compiled
LabView program. Since LabView implements its own controls, the Microsoft Window standard routines for finding controls
cannot be used. Instead, the relative position of each control must be know in pixel coordinates.

# Status
This program just started to work, but is already astonishingly stable in internal tests. Feel free to try it out
yourself. If you run into problems or have questions, open an issue or write me a message.