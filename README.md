# powerpoint-autohotkey

Helper keys for using powerpoint with logitech style clickers

This repository contains four scripts for use with AutoHotkey, which you can download and install by visiting [www.autohotkey.com](https://www.autohotkey.com/)

Background
---

The start/stop key in a logitech remote alternates between ```Esc``` and ```F5```.

The blackout key (to the right of the start/stop key) sends the OEM_PERIOD key ( virtualKey 0xBE, aka the "." key)

We trap these keys and replace them or ignore them to provide a more stable experience for the presenter.


PPT - Disable Logitech Blackout and Escape.ahk & PPT - Disable Logitech Blackout Key.ahk
===

![image](https://github.com/jonathan-annett/powerpoint-autohotkey/blob/c8f7bb48a84fda4a0b1bf49dbfd33902ac61cb8a/Ignore%20Blackout.png)

* `PPT - Disable Logitech Blackout Key.ahk` may require the "Start Slideshow" key to be pressed twice, as the default behaviour for the clicker is to alternate between F5 and Escape


PPT - Notes Scroller.ahk
===

![image](https://github.com/jonathan-annett/powerpoint-autohotkey/blob/24269185b5a01e2eecdb6220af85b3fc9cd09f08/Notes%20Scroller.png)


PPT - Logitech Toggle Laser.ahk
===

Replaces the OEM_PERIOD key (blackout) with Control L when in PowerPoint.

This toggles the laser pointer in PowerPoint

Keyboard Helpers
===
All scripts allow you to start and stop the presentation, and scroll the notes view using the keyboard

![image](https://github.com/jonathan-annett/powerpoint-autohotkey/blob/0df841ba6b2bcd6a340976e2629aac80e9a7424d/key%20scroll.png)
