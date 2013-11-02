# Python automatisation wrapper around the NanoWrite software by nanoscribe GmbH
## Introduction
The NanoWrite software by nanoscribe GmbH is the user interface to their awesome 3D printer.
On its own, this software provides an intuitive user interface and comes with its own simple scripting language,
called GWL.

For most users, this will probably be enough already. But there are times where one needs to automate specific
processes like finding a specific marker, doing optical auto-focus or load a GWL file which was generated on the fly.

## Features
This project implements a simple Python class to export the following features:

1. Get the current log.

   This includes parsing the timestamps and correctly evaluating multiple line log entries.
2. Execute GWL commands.
    + In the mini GWL window.
    + Load a specific GLW (structure) file
    + Load a set of GLW files and read back their results, such a images etc.
3. Switch to camera view.
4. Get piezo and stage coordinates.
5. Read current progress time and the time estimate of the structure.
6. Check if the current job have finished. (This is important and not so easily implemented as it sounds.)
7. Abort the current job.
8. Get the current camera picture as binary file.
9. Correctly read, set and handle the z-axis inversion feature.

Additionally, **this project includes a sample XML-RPC server which exposes exactly the same features and commands.
So you do not need to program in Python to use this code**. A list of XML-RPC client implementations can be found
[here](https://en.wikipedia.org/wiki/XML-RPC#Implementations).

## Technical implementation
This wrapper automates the nanowrite software by simulating series of mouse and keyboard presses. Great care has
been taken to make the process as stable as possible. Matters are further complicated by NanoWrite being a compiled
LabView program. Since LabView implements its own controls, the Microsoft Window standard routines for finding controls
cannot be used. Instead, the relative position of each control must be know in pixel coordinates.

# Status
This program just started to work, but is already astonishingly stable in internal tests. Feel free to try it out
yourself. If you run into problems or have questions, open an issue or write me a message.