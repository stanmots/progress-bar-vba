A VBA Progress Bar for the Microsoft Office Applications
=======================================================

If you want to keep your users informed how the code is progressing, you should definitely find useful this piece of software. 

For example, you can import this progress bar directly into your Excel workbook and use it very easily with all your favorite macros.

###System Requirements:
- Microsoft Office Applications **only**.

**Note:** It is incompatible with the Microsoft Office for Mac.

Basic Usage:
-----------

Let's have a look at the progress bar image:

![Progress Bar Image][ProgressBarImageId]


Here are the main steps you must complete to use this progress bar with your project:

- Import all the source files from this repo. Typically, you do this in the VBE (Visual Basic Editor) by right-clicking on the project's name and selecting 'import' option;
- Change almost every property with a single code line as you need.

**Note:** You also need the vba userform binary file with the 'frx' extension. It is included in the *Release* package.

This is all coding you need for the typical case:

~~~ vbnet
Public Sub TestProgressBarForm()

'show progress bar
ProgressBarForm.Show

'change the main properties
ProgressBarForm.SetCurrentOperationLabelText "Current Operation Title"
ProgressBarForm.SetMainLabelText "Main Title"

'add information about the current operation, time will be added automatically
ProgressBarForm.AddMessageToDetailsBox "Program started..."

'change progress bar indicator
ProgressBarForm.IncreaseProgressByPercent 50

'...some very smart code must be here...

'now the progress will be 100
ProgressBarForm.IncreaseProgressByPercent 50

ProgressBarForm.AddMessageToDetailsBox "All operations have been finished!"

End Sub
~~~

If you wanna use it inside loops here is a convenient method:
~~~ vbnet
'firstly, set loops parameters (overall percentage you wanna add, loops number) 
ProgressBarForm.SetLoopsParameters 100, 10
Dim i As Integer
For i = 0 To 9
'now you can freely call this method each iteration
ProgressBarForm.IncreaseProgressInsideLoop
Next

~~~

####That's all. Thanks for your attention!

**Note:** The project is open source under the [MIT License (MIT)](https://github.com/storix/progress-bar-vba/blob/master/LICENSE.md).

[ProgressBarImageId]: http://s14.postimg.org/70zjqr71d/Screen_Shot_2015_01_16_at_10_54_50_PM.png  "Progress Bar"
