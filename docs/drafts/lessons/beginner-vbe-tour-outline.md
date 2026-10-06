# Lesson outline — Tour the Visual Basic Editor

**Publish into:** [src/content/lessons/beginner-vbe-tour.md](../../src/content/lessons/beginner-vbe-tour.md)

```yaml
title: Tour the Visual Basic Editor
track: beginner
order: 1
summary: Open the VBE, understand the Project Explorer, and run your first Sub.
draft: false
```

**Target:** 900–1,200 words | **Code blocks:** 2–3 | **Template:** [beginner-first-macro.md](../../src/content/lessons/beginner-first-macro.md)

---

## Goal

Get comfortable inside the Visual Basic Editor (VBE) so you can write and run code instead of recording fragile macros.

## When to use this

The VBE is used whenever a macro needs to be edited or updated, or when you want to create repeatable automations.

## Open the VBE

There are two ways to get into the Visual Basic Editor.
1. Using the ribbon
    a. Click on the Developer tab, then the Visual Basic button
    ![Developer tab with Visual Basic button](/images/lessons/beginner-vbe-tour/developer-ribbon-0.webp)
    b. If you don't see the Developer tab,  right click on the ribbon, select 'Customize the Ribbon' and click the check box next to Developer to enable it
    ![Ribbon right click to customize](/images/lessons/beginner-vbe-tour/developer-tab-right-click.webp)
    ![Customize Ribbon window](/images/lessons/beginner-vbe-tour/developer-tab-2.webp)
2. Using Hotkeys (My personal favorite since it works regardless of the ribbon)
    a. Alt + F11


## The basic window explained

When the VBE opens, you will see the editor with the menubar/toolbar and two windows on the left. The two windows on the left are the Project Explorer and the Property window.
![Main Visual Basic Editor Window](/images/lessons/beginner-vbe-tour/editor-1.webp)

The toolbar has the standard toolbar that includes some useful features, but I like to include the Edit and Debug toolbars, too. To add them in the Menu Bar click View > Toolbars or right click on the toolbar to select the Edit and Debug toolbars.
![Visual Basic Editor Toolbars Context Menu](/images/lessons/beginner-vbe-tour/editor-toolbars-1.webp)
![Visual Basic Editor Toolbars](/images/lessons/beginner-vbe-tour/editor-toolbars-2.webp)

The Project Explorer is the working tree of the objects in the Excel file.  In our practice workbook, you will see a folder that houses the workbook and worksheets.  

The Properties window shows the different things an object has, for example, a sheet has items like Name and Visible which can be used in our code.
![Main Visual Basic Editor Window](/images/lessons/beginner-vbe-tour/editor-2.webp)

We will dive into these other objects in later lessons, for now, let's insert a module.

## Insert a standard module

There are multiple ways to insert a module, the tool bar, the icon or right click in the project explorer.
![Insert a module - Method 1](/images/lessons/beginner-vbe-tour/insert-module-1.webp)
![Insert a module - Method 2](/images/lessons/beginner-vbe-tour/insert-module-2.webp)
![Insert a module - Method 3](/images/lessons/beginner-vbe-tour/insert-module-3.webp)

When you insert a module, a new folder will be displayed in the Project Explorer named 'Modules' with an object 'Module1'.

When you insert a module, it will open a window in the rest of VBE.  This window can be resized using the icons in the top right 
![Visual Basic Editor Window with Code Window](/images/lessons/beginner-vbe-tour/editor.webp)

If the Code window doesn't open, double click on Module1 in the Project Explorer window.

## Write and run Hello Sub

Let's "write" your first macro, click in the window and you will see the blinking cursor to type.

Now just copy and paste the code below into the newly inserted module.

```vb
Option Explicit

Sub HelloDecodingVBA()
    MsgBox "VBE is open. You are ready to write code.", vbInformation, "Decoding VBA"
End Sub
```
Click in the sub and press F5 or click the play button in the toolbar.
![Running a Macro](/images/lessons/beginner-vbe-tour/run-macro.webp)

If you want to step through the code, Press F8 instead.

## Save as macro-enabled

Once you add a macro into a workbook, it will need to be saved as a Macro enabled workbook, .xlsm, to keep the macros.  If you save it as a .xlsx, it will remove all the macros from the workbook.
![Saving a Macro-enabled workbook](/images/lessons/beginner-vbe-tour/save-as-xlsm.webp)


## What to notice

- In the Project Explorer, reusable macros live under Modules (for example, Module1), not under Microsoft Excel Objects sheet nodes unless you deliberately use sheet code.
- Your macro is a Sub … End Sub procedure; click inside HelloDecodingVBA before you run so VBA knows which procedure to execute.
- F5 runs the entire Sub; F8 steps line by line—useful when you want to see each line run once.
- Macros are part of the workbook file—save as .xlsm (macro-enabled) so the code stays with the file (plain .xlsx does not keep VBA).

## Common mistakes

- Excel crashes happen and sometimes you can lose work if you haven't saved in a while; so save often.
- Saving as `.xlsx` instead of `.xlsm` (macros are stripped).
- Writing code in the worksheet module when a standard module is clearer for reusable Subs.
- Overwriting production data while debugging or testing.

## Practice

Some things to practice include changing the message in the MsgBox, creating a second Sub, and displaying the ActiveSheet Name in the MsgBox.

```vb
    MsgBox ActiveSheet.name, vbInformation, "Decoding VBA"
```
[DecodingVBA-Lesson01-VBE-Tour.xlsm](/downloads/DecodingVBA-Lesson01-VBE-Tour.xlsm)

## Next up

Next, we will look into Option Explicit and why it is important to include in your code.
<!-- LINK: Option Explicit lesson — beginner-option-explicit -->
