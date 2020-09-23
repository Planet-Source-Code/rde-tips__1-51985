<div align="center">

## Tips

<img src="PIC2004225717305514.gif">
</div>

### Description

Some very basic but handy tips for newbies (and maybe not so newbies). Hope you learn something.

I've added quite a bit more to the article, mostly very basic info on the vb language, but we all had to learn this stuff when we started out.
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[Rde](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/rde.md)
**Level**          |Beginner
**User Rating**    |3.8 (19 globes from 5 users)
**Compatibility**  |VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0
**Category**       |[Miscellaneous](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/miscellaneous__1-1.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/rde-tips__1-51985/archive/master.zip)





### Source Code

<font face="Verdana" size="-1">&#160;
<h3 align="center">A Brief History of Basic</h3>
<ul>
	<li> BASIC - Beginner's All-purpose Symbolic Instruction Code. </li>
	<li> This language was developed in the early 1960's at Dartmouth College. </li>
	<li> Answer to complicated programming languages (FORTRAN, Algol, Cobol...). First timeshare language. </li>
	<li> In the mid-1970's, two college students write first Basic for a microcomputer (Altair) - cost $350 on cassette tape. You may have heard of them: Bill Gates and Paul Allen. </li>
	<li> Every Basic since then is essentially based on that early version. Examples include: GW-Basic, QBasic, QuickBasic. </li>
	<li> Visual Basic was introduced in 1991. </li>
</ul>
<hr width="95%" size="2" align="center" />
<p><b>Tips</b></p>
<p>In VB's Tools > Options > Editor Format tab > Code Colors<br />
Select Identifier Text > then Foreground > Dark Red!</p>
<hr width="95%" size="2" align="center" />
<p>I recommend that you uncheck:</p>
<p>Tools > Options > Editor tab > Auto Syntax Check</p>
<p>Syntax checking is still enabled but instead of freezing the editor until you dismiss the msgbox and correct the error, the offending code is simply displayed in red, allowing you to continue coding unhindered, and can return to the 'red devil' when finnished what you're doing. Rd.</p>
<hr width="95%" size="2" align="center" />
<p>I often see this:</p>
<pre>Dim var1, var2, var3 As String</pre>
<p>This initializes 2 variants and 1 string, it's the same as this:</p>
<pre>Dim var1 As Variant, var2 As Variant, var3 As String</pre>
<p>I assume the coder desires:</p>
<pre>Dim var1 As String, var2 As String, var3 As String</pre>
<p>Use VarType or TypeName to see for yourself.</p>
<p>Added this bit thanks to Timothy Marin. Some say not to use these (and james kahl has a point) but it's up to you:</p>
<pre>
Dim fff$ ' Declares a String variable
Dim hhh% ' Declares a Integer variable
Dim ggg& ' Declares a Long variable
Dim iii! ' Declares a Single variable
Dim jjj# ' Declares a Double variable
Dim kkk@ ' Declares a Currency variable</pre>
<hr width="95%" size="2" align="center" />
<p>I often see this:</p>
<pre>
If Dir$(sFileName) <> "" Then
  ' I believe I have a file,
  ' do some work on it
End If</pre>
<p>This code will fail if sFileName = "" because Dir will return the name of the first file found (with any name) You can do this instead:</p>
<pre>
If sFileName <> "" Then
  If Dir$(sFileName) <> "" Then
    ' We have a file, do some work on it
  End If
End If</pre>
<p>Or as Luke H so rightly pointed out, you should do this:</p>
<pre>
If LenB(sFileName) <> 0 Then
  If LenB(Dir$(sFileName)) <> 0 Then
    ' We have a file, do some work on it
  End If
End If</pre>
<p>This lot's from Bruce McKinney's book 'Hardcore VB'</p>
<p>That statement works until you specify a file on an empty floppy or CD-ROM drive. Then you’re stuck in a message box. Here’s another common one:</p>
<pre>
fExist = FileLen(sFullPath)</pre>
<p>It fails on 0-length files — uncommon but certainly not unheard of. My theory is that the only reliable way to check for file existence in VB (without benefit of API calls) is to use error trapping. I’ve challenged many Visual Basic programmers to give me an alternative, but so far no joy. Here’s the shortest way I know:</p>
<pre>
Function FileExists(sSpec As String) As Integer
  On Error Resume Next
  Call FileLen(sSpec)
  FileExists = (Err = 0)
End Function</pre>
<p>This can’t be very efficient. Error trapping is designed to be fast for the no fail case, but this function is as likely to hit errors as not. Perhaps you’ll be the one to send me a Basic-only ExistFile function with no error trapping that I can’t break. Until then, here’s an API alternative:</p>
<pre>
Function ExistFileDir(sSpec As String) As Boolean
  Dim af As Long
  af = GetFileAttributes(sSpec)
  ExistFileDir = (af <> -1)
End Function</pre>
<p>I didn’t think there would be any way to break this one, but it turns out that certain filenames containing control characters are legal on Windows 95 but illegal on Windows NT. Or is it the other way around? Anyway, I have seen this function fail in situations too obscure to describe here. Bruce McKinney.</p>
<hr width="75%" size="1" align="center" />
<p>Here's my solution:</p>
<pre>
Function FileExists(sFileSpec As String) As Boolean
  If (sFileSpec = vbNullString) Then Err.Raise 5
  On Error GoTo NoGo
  Dim Attribs As Long
  Attribs = FileSystem.GetAttr(sFileSpec)
  If (Attribs <> -1) Then
    FileExists = ((Attribs And vbDirectory) <> vbDirectory)
  End If
NoGo:
End Function</pre>
<pre>
Function DirExists(sPath As String) As Boolean
  If (sPath = vbNullString) Then Err.Raise 5
  On Error GoTo NoGo
  Dim Attribs As Long
  Attribs = FileSystem.GetAttr(sPath)
  If (Attribs <> -1) Then
    DirExists = ((Attribs And vbDirectory) = vbDirectory)
  End If
NoGo:
End Function</pre>
<hr width="95%" size="2" align="center" />
<h3 align="center">Data types in VB</h3>
<p>Numeric Data : Integers (whole numbers without decimal
places) and Real (decimals, or floating-point numbers).</p>
<table border="1" cellspacing="0" cellpadding="4">
	<tr>
		<th>
				Suf
		</th>
		<th>
				Type
		</th>
		<th>
				Storage
		</th>
		<th>
				Range
		</th>
	</tr>
	<tr>
		<td >&#160;
		</td>
		<td >
			Byte
		</td>
		<td >
			1 byte
		</td>
		<td >
			&#160;0 to 255
		</td>
	</tr>
	<tr>
		<td >
			<b>
				%
			</b>
		</td>
		<td >
			Integer
		</td>
		<td >
			2 bytes
		</td>
		<td >
			-32,768 to 32,767
		</td>
	</tr>
	<tr>
		<td >
			<b>
				&
			</b>
		</td>
		<td >
			Long
		</td>
		<td >
			4 bytes
		</td>
		<td >
			-2,147,483,648 to 2,147,483,647
		</td>
	</tr>
	<tr>
		<td >
			<b>
				!
			</b>
		</td>
		<td >
			Single
		</td>
		<td >
			4 bytes
		</td>
		<td >
			-3.42823E+38 to -1.401298E-45 (neg) and<br />
    &#160;1.401298E-45 to 3.42823E+38 (pos)
		</td>
	</tr>
	<tr>
		<td >
			<b>
				#
			</b>
		</td>
		<td >
			Double
		</td>
		<td >
			8 bytes
		</td>
		<td >
			-1.79769313486232E+308 to<br />
    -4.94065645841247E-324 (negative) and<br />
    &#160;4.94065645841247E-324 to<br />
    &#160;1.79769313486232E+308 (positive)
		</td>
	</tr>
	<tr>
		<td >
			<b>
				@
			</b>
		</td>
		<td >
			Currency
		</td>
		<td >
			8 bytes
		</td>
		<td >
			-922,337,203,685,477.5808 to<br />
    &#160;922,337,203,685,477.5807 (the extra<br />
    precision ensures accuracy to 2 dec places)
		</td>
	</tr>
	<tr>
		<td >&#160;
		</td>
		<td >
			Decimal
		</td>
		<td >
			12 bytes
		</td>
		<td >
     +/-79,228,162,514,264,337,593,543,950,335<br />
     (with no decimal, or up to 28 decimal places)<br />
     +/-7.9228162514264337593543950335
		</td>
	</tr>
</table>
<h3>Shift State</h3>
<table border="0" cellspacing="2" cellpadding="4">
	<tr>
		<th align="left">
			Constant
		</th>
		<th align="center">
			Value
		</th>
		<th align="left">
			Description
		</th>
	</tr>
	<tr>
		<td align="left">
			vbShiftMask
		</td>
		<td align="center">
			1
		</td>
		<td align="left">
			SHIFT key bit mask.
		</td>
	</tr>
	<tr>
		<td align="left">
			vbCtrlMask
		</td>
		<td align="center">
			2
		</td>
		<td align="left">
			CTRL key bit mask.
		</td>
	</tr>
	<tr>
		<td align="left">
			vbAltMask
		</td>
		<td align="center">
			4
		</td>
		<td align="left">
			ALT key bit mask.
		</td>
	</tr>
</table>
<p>Presently, only three of the 32 bits in the Shift parameter<br />
are used. In future versions of Visual Basic, however, these<br />
other bits may be used. Therefore, as a precaution against<br />
future problems, you should mask these values appropriately<br />
before performing any comparisons. Use a bitwise And to mask<br />
the Shift parameter:</p>
<pre>
Dim ShiftState As Integer
ShiftState = Shift And vbShiftMask</pre>
<p>In the above example ShiftState will hold zero if the Shift<br />
key was not pressed, or one if pressed (in any combination<br />
with the Ctrl and Alt keys). Likewise, you can mask the Shift<br />
parameter against vbCtrlMask to return zero or two, and<br />
vbAltMask to return zero or four.</p>
<pre>
Dim ShiftDown As Boolean
Dim CtrlDown As Boolean
Dim AltDown As Boolean</pre>
<pre>
ShiftDown = (Shift And vbShiftMask) = vbShiftMask
CtrlDown = (Shift And vbCtrlMask) = vbCtrlMask
AltDown = (Shift And vbAltMask) = vbAltMask</pre>
<hr width="95%" size="2" align="center" />
<p>True-conditions perform faster. So, if you can make assumptions about
your conditions, set up the code so that the test returns True.</p>
<hr width="95%" size="2" align="center" />
<p>Before continuing, SAVE your project to disk for safety!</p>
<hr width="95%" size="2" align="center" />
<h3 align="center">Max Path Length</h3>
<pre>Const MAX_PATH As Long = 260</pre>
<p>The maximum length, in characters, of a file path supported by the
specified file system. A filename component is actually that portion
of a file path between backslashes.</p>
<p>Under NT (Intel) and Win95 it can be up to 259 (MAX_PATH - 1) characters
long. This length must include the drive, path, filename, commandline
arguments and quotes (if the string is quoted).</p>
<p>Notice that the MAX_PATH constant is assigned 260 on Windows 9x systems.
This is because it combines the root ("x:\"), the Maximum Component Length
value (255), plus a possible trailing backslash ("\") character.</p>
<pre>Len(sPath) <= 3 + 255 + 1  or  Len(sPath) < MAX_PATH</pre>
<p>The complete path <i><b>must be less than</b></i> MAX_PATH characters.</p>
<hr width="95%" size="2" align="center" />
<h3 align="center">Logical Operators</h3>
<p>The logical operators enable you to combine two or more sets of
conditional comparisons.</p>
<table border="0" cellspacing="2" cellpadding="4">
	<tr>
		<th>
			And
		</th>
		<td>
			Both sides must be True (to return True)
		</td>
	<tr>
	</tr>
		<th>
			Or
		</th>
		<td>
			Only one side need be True, or both
		</td>
	</tr>
	<tr>
		<th>
			Xor
		</th>
		<td>
			Only one side must be True, not both
		</td>
	<tr>
	</tr>
		<th>
			Not
		</th>
		<td>
			Reverses (inverts) boolean condition
		</td>
	</tr>
</table>
<p>The <b>And</b> logical operator requires both sides to be True to
return True.</p>
<pre>
 If (x >= 1) And (x <= 10) Then ...</pre>
<p>The <b>Or</b> logical operator needs only one side to be True to
return True. This operator is really an Inclusive Or.</p>
<pre>
 If (y = 0) Or (z <> 10) Then ...</pre>
<p>The <b>Xor</b> logical operator requires that only one side CAN be
True to return True. Therefore, its is an Exclusive Or.</p>
<pre>
 If (count1 = limit) Xor (count2 = limit) Then
   CountSyncErrorOccured
 End If</pre>
<p>The <b>Not</b> logical operator inverts the boolean value.</p>
<p>The following two code examples both reverse the value:</p>
<pre>
 result = Not (expression)
 result = (expression) Xor True</pre>
<p>Be careful with the Xor and Not operators, as they only work
(as you might expect) with boolean True and False values. So
expression must evaluate to a boolean True or False value.</p>
<p>Note - True equates to -1, and False equates to 0.</p>
<p>Because zero equates to False, and all other numbers equate
to True when tested within a conditional (an If statement for
example) you can generally do this:</p>
<pre>
 If iNum Then
   'Do something
 End If</pre>
<p>If iNum is not zero it will equate to True, including
negative values:</p>
<pre>
 3 = True
 2 = True
 1 = True
 0 = False
 -1 = True
 -2 = True
 -3 = True</pre>
<p>But if you wanted to reverse the condition as follows,
it may not work as you expect:</p>
<pre>
 If Not iNum Then
   'Do something
 End If</pre>
<p>Only if iNum is -1 will the conditional equate to False.
Any other value including zero will equate to True:</p>
<pre>
 Not 3 = -4 ' True
 Not 2 = -3 ' True
 Not 1 = -2 ' True
 Not 0 = -1 ' True (Not False)
 Not -1 = 0 ' False (Not True)
 Not -2 = 1 ' True
 Not -3 = 2 ' True</pre>
<p>So do the following when using Not with numeric values:</p>
<pre>
 If Not CBool(iNum) Then
  'Do something
 End If</pre>
<pre>
 Not CBool(3) = False ' Not True
 Not CBool(2) = False ' Not True
 Not CBool(1) = False ' Not True
 Not CBool(0) = True ' Not False
 Not CBool(-1) = False ' Not True
 Not CBool(-2) = False ' Not True
 Not CBool(-3) = False ' Not True</pre>
<p>Note that Xor works the same.</p>
<hr width="95%" size="2" align="center" />
<h3 align="center">Conditional Operators</h3>
<p>VB supports six conditional operators:</p>
<pre>
 =       Equal to
 >       Greater than
 <       Less than
 >=      Greater than or equal to
 <=      Less than or equal to
 <>      Not equal to</pre>
<p>VB also supports a special kind of conditional operator:</p>
<pre>
 Like     Performs comparisons using wildcards</pre>
<p>Here are the widcards that can be used with Like:</p>
<pre>
 *       Any character or characters
 ?       Any alpha character (letters)
 #       Any numeric character (numbers)
 []      Encloses possible characters
 -       Specifies a range</pre>
<p>e.g:</p>
<pre>
 "This string" Like "This*"       returns True
 "This string" Like "This ???ing"    returns True
 "Numeric 123" Like "Numeric ###"    returns True
 "Version 2 b" Like "Version [123] *"  returns True
 ""      Like "[]"        returns True
 "E"      Like "[C-H]"       returns True</pre>
<p>Use the [] to test for a possible character within a group.</p>
<p>By using a hyphen (–) to separate the upper and lower bounds
of the range, charlist can specify a range of characters. The
meaning of a specified range depends on the character ordering
valid at run time (as determined by Option Compare and the
locale setting of the system the code is running on). Using
the Option Compare Binary, the range [A–E] matches A, B, C, D, E.
With Option Compare Text, [A–E] matches A, a, À, à, B, b,... E, e.
The range does not match Ê or ê because accented characters
fall after unaccented characters in the sort order.</p>
<h3 align="center">Mathematical Operators</h3>
<p>The mathematical operators perform calculations on numerical
values:</p>
<pre>
 ()      Parenthesis
 ^       Exponentiation (Power Of)
 *       Multiplication
 /       Division
 \       Integer Division
 Mod      Modulus
 +       Addition
 -       Subtraction</pre>
<h3 align="center">Operator Precedence</h3>
<p>Precedence is the order of importance given to operators in VB.
In other words, precedence determines which part of an expression
will be executed first.</p>
<p>The following is the order of precedence from highest to lowest:</p>
<pre>
 ()      Parenthesis
 ^       Exponentiation (Power Of)
 * / \ Mod Multiplication, Division, Int Division and Modulus
 + -     Addition and Subtraction
 Like     Performs comparisons using wildcards
 Not      Reverses (negates) boolean condition
 And      Both sides must be true
 Or      Only one side need be true, or both
 Xor      One side must be true, but not both</pre>
<hr width="95%" size="2" align="center" />
<p>Before continuing, SAVE your project to disk for safety!</p>
<hr width="95%" size="2" align="center" />
<p>Byte arrays are the only way to store binary data in a stable
format that won't be modified by Unicode conversion.</p>
<hr width="95%" size="2" align="center" />
<p>To Add a text file or other non-standard file into a VB project
edit the .vbp file and insert a line similar to this:</p>
<pre>RelatedDoc=readme.txt</pre>
<p>The file will be displayed in the Resources section of the
project explorer.</p>
<hr width="95%" size="2" align="center" />
<p>A loop used to remove selected items from a list without error:</p>
<pre>
For i = lstData.ListCount - 1 To 0 Step -1
  If lstData.Selected(i) Then lstData.RemoveItem i
Next</pre>
<hr width="95%" size="2" align="center" />
<p>Option buttons use the property .Value = True|False while
Checkboxes use the property .Value = 0|1|2 to specify checked.</p>
<hr width="95%" size="2" align="center" />
<p>You can put more than one statement on a line by separating them
with a colon <b>:</b></p>
<pre>Dim myInt%: myInt = 0</pre>
<p>Thanks Timothy Marin for this tip.</p>
<hr width="95%" size="2" align="center" />
<p>Before continuing, SAVE your project to disk for safety!</p>
<hr width="95%" size="2" align="center" />
<p>If you want to place controls in a frame but already have the controls
on the form, just select all controls (with selection tool or hold down
the CTRL key as you click each control), and CUT, then place the
frame on the form, then with the frame selected, PASTE.</p>
<hr width="95%" size="2" align="center" />
<h3 align="center">KeyPress and KeyDown Events</h3>
<p>The KeyPress event occurs when the user presses the Uppercase and
Lowercase letters, Numeric digits, Punctuation keys, and the Enter,
Tab, and Backspace keys.</p>
<p>Some VB Constants are: vbKeyReturn, vbKeyTab and vbKeyBack.</p>
<p>KeyPress events capture just the main ASCII characters (letters,
numbers and punctuation) plus Backspace, Enter and Tab.</p>
<p>KeyPress handles the shift state itself, passing the event procedure
the correct code. The event object recieves the keycode (as was or
modified) AFTER the procedure ends.</p>
<p>This makes KeyPress the event handler to use when you wish to process
and/or modify the characters before they are displayed in the form's
event object.</p>
<p>KeyDown events capture ALL keyboard keys, but only recognize letters
in all caps, so you must test for shift state as well for lowercase.
KeyUp (like KeyDown) receives only uppercase keycodes.</p>
<p>KeyDown does not wait to pass the key code on to the form object,
while KeyPress passes the key code on to the event object after
the procedure ends.</p>
<p>With KeyDown the event object handles the shift state itself, so
the event object receives upper or lowercase according to the shift
and caps lock key states - then the KeyDown event procedure runs.</p>
<hr width="95%" size="2" align="center" />
<p>You might think that you could save space by declaring a
variable As Byte or As Integer instead of As Long. However, on
32-bit operating systems the code to load a Long is faster and
more compact than the code to load shorter data types.</p>
<p>Not only could the extra code exceed the space saved, but there
might not be any space saved to begin with — because of alignment
requirements (32-bit) for modules and data.</p>
<hr width="95%" size="2" align="center" />
<h3 align="center">Subs and Funcs</h3>
<p>A subroutine does not return a value.</p>
<pre>
Private Sub cmdSubCalculate_Click()
  Call multiply1(2, 3)
End Sub</pre>
<pre>
Private Sub multiply1(ByVal x As Integer, ByVal y As Integer)
  Dim z As Integer
  z = x * y
  txtResult.Text = z
End Sub</pre>
<p>A function returns a value by assigning the resulting<br />
value of its processing to a 'variable' (the name of the<br />
function) which is returned to the calling subroutine or<br />
function.</p>
<pre>
Private Sub cmdFuncCalculate_Click()
  txtResult.Text = Multiply2(2, 3)
End Sub</pre>
<pre>
Private Function Multiply2(ByVal x As Integer, ByVal y As Integer) As Integer
  Dim z As Integer
  z = x * y
  Multiply2 = z
End Function</pre>
<p>Both subs and functions can have arguments passed by<br />
reference, allowing the source variable to be modified<br />
by the procedure.</p>
<pre>
Private Sub cmdByRefCalculate_Click()
  Dim z As Integer
  If (Multiply3(2, 3, z) Then
    txtResult.Text = z
  End If
End Sub</pre>
<pre>
Private Function Multiply3(ByVal x As Integer, ByVal y As Integer, ByRef z As Integer) As Boolean
  z = x * y
  Multiply3 = True
End Function</pre>
<hr width="95%" size="2" align="center" />
<p><b>You can add quit confirmation to a windows standard exit methods:</b></p>
<pre>
Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
  ' A forms QueryUnload event occurs immediately before the
  ' form unloads.
  ' UnloadMode is zero when a window is closed by using any
  ' of the standard exit methods (by clicking the [x] close button,
  ' by selecting close from the windows context menu, by pressing
  ' ALT-F4, or by double-clicking the window icon in the top-left
  ' corner).
  ' Cancel (passed by reference to the event) is zero (False), and
  ' so the form will unload; False means 'not to Cancel the unload'.
  ' You can cancel the unloading of the form by setting Cancel to
  ' one (True); so saying 'yes to Cancel the unload'.
  If UnloadMode = 0 Then
    Dim dialogtype As Integer
    Dim title, msg As String
    Dim response As Integer
    dialogtype = vbYesNo + vbQuestion
    title = "Name of program"
    msg = "Are you sure?"
    response = MsgBox(msg, dialogtype, title)
    If response = vbNo Then
      Cancel = True
    End If
  End If
End Sub</pre>
<hr width="75%" size="1" align="center" />
<p>The UnloadMode variable in the Query_Unload event indicates how
this event was triggered by containing one of the five values in
the following table.</p>
<pre>QueryUnloadConstants:
 vbFormControlMenu = 0
   The user chose the Close command on the Control-menu box.
 vbFormCode = 1
   The application used the Query_Unload method itself.
 vbAppWindows = 2
   The operating system is being shut down, or the user is
   logging off.
 vbAppTaskManager = 3
   The application is being shut down by the Task Manager.
 vbFormMDIForm = 4
   An MDI form, which closes all child forms belonging to it,
   is being closed.
 vbFormOwner = 5
   The owner of the form is closing.</pre>
<hr width="95%" size="2" align="center" />
<h3 align="center">Select Case Statement</h3>
<p>Executes one of several groups of statements, depending
on the value of an expression.</p>
<pre>
Select Case Index
   Case 0
    Grade = "first"
   Case 1
    Grade = "second"
   Case 2
    Grade = "third"
   Case 3
    Grade = "fourth"
   Case 4
    Grade = "fifth"
   Case 5
    Grade = "sixth"
End Select</pre>
<p>The same Case Statement using colons:</p>
<pre>
Select Case Index
  Case 0 :  Grade = "first"
  Case 1 :  Grade = "second"
  Case 2 :  Grade = "third"
  Case 3 :  Grade = "fourth"
  Case 4 :  Grade = "fifth"
  Case 5 :  Grade = "sixth"
End Select</pre>
<p>VB also offers a way of testing for a condition
with the Is keyword added to Case:</p>
<pre>
Select Case testscore
  Case Is >= 80
    student_grade = "A"
  Case Is >= 65
    student_grade = "B"
  Case Is >= 50
    student_grade = "C"
  Case Else
    student_grade = "F"
End Select</pre>
<p>So the following two examples are the same:</p>
<pre>
Select Case Format(today, "mmmm")
  Case "January":     optjan.Value = True
  Case "February":    optfeb.Value = True
  Case "March":      optmar.Value = True
  Case "April":      optapr.Value = True
  Case "May":       optmay.Value = True
  Case "June":      optjun.Value = True
  Case "July":      optjul.Value = True
  Case "August":     optaug.Value = True
  Case "September":    optsep.Value = True
  Case "October":     optoct.Value = True
  Case "November":    optnov.Value = True
  Case "December":    optdec.Value = True
End Select</pre>
<pre>
Select Case Format(today, "mmmm")
  Case Is = "January":  optjan.Value = True
  Case Is = "February":  optfeb.Value = True
  Case Is = "March":   optmar.Value = True
  Case Is = "April":   optapr.Value = True
  Case Is = "May":    optmay.Value = True
  Case Is = "June":    optjun.Value = True
  Case Is = "July":    optjul.Value = True
  Case Is = "August":   optaug.Value = True
  Case Is = "September": optsep.Value = True
  Case Is = "October":  optoct.Value = True
  Case Is = "November":  optnov.Value = True
  Case Is = "December":  optdec.Value = True
End Select</pre>
<p>The usage for the Case Is format allows the testing of conditions
that don't have to be an exact match (=), but can also be other
conditions (>, <, >=, <=, etc). Only the use of simple comparisons
are allowed, so no logical operators (And, Or, Xor, or Not) can
be used.</p>
<pre>
Select Case varInteger
  Case Is <= 150:  MsgBox "Is 150 or less"
  Case Is <= 200:  MsgBox "Is between 151 and 200 inclusive"
  Case Else:    MsgBox "Is 201 or greater"
End Select</pre>
<p>It is important to realize that with such conditional tests the order
of each Case Is matters. Consider if the above Case statement was
like this:</p>
<pre>
Select Case varInteger
  Case Is <= 200:  MsgBox "Is 200 or less"
  Case Is <= 150:  MsgBox "Is 150 or less"
  Case Else:    MsgBox "Is 201 or greater"
End Select</pre>
<p>Even if the integer value was below 150 the condition would still
execute only the code corresponding to the first Case tested.</p>
<p>In addition to the formats used above, Select Case statements can
also include the To keyword, as follows:</p>
<pre>
Select Case Asc(Char)
  Case 65 To 90: MsgBox "Uppercase 'A' to 'Z' inclusive"
  Case 97 To 122: MsgBox "Lowercase 'a' to 'z' inclusive"
End Select</pre>
<p>You can combine the different formats into a single Case statement:</p>
<pre>
Select Case varInteger
  Case 100 To 130, 140
    MsgBox "Is between 100 and 130 inclusive, or is 140"
  Case 150 To 180, 190, Is >= 200
    MsgBox "Is between 150 and 180 inclusive, or is 190, or is 200 or greater"
  Case Else
    MsgBox "All else (below 100, 131 to 139, etc)"
End Select</pre>
<p>You also can specify ranges and multiple expressions for
character strings.</p>
<p>In the following example, Case matches strings that are exactly
equal to everything, strings that fall between nuts and soup in
alphabetic order, and the current value of TestItem:</p>
<pre>
  Case "everything", "nuts" To "soup", TestItem</pre>
<hr width="95%" size="2" align="center" />
<h3 align="center">Gaussian Rounding</h3>
<p>A remark by René Rhéaume, 21.09.01</p>
<p>This "Banker's" method uses the Gauss rule that if you are
in an perfect half case, you must round to the nereast digit
that can be divided by 2 (0,2,4,6,8). This rule is important
to obtain more accurate results with rounded numbers after
operation.</p>
<p>Now, an example :</p>
<pre>
       2 digits        2 digits
Unrounded  "Standard" rounding  "Gaussian" rounding
 54.1754   54.18         54.18
 343.2050   343.21         343.20
+106.2038  +106.20        +106.20
=========  =======        =======
 503.5842   503.59         503.58</pre>
<p>Which one is nearer from unrounded result? The "Gaussian" one
(Difference of 0.0042 with "Gaussian/Banker" and 0.0058
with "Standard" rounding.)</p>
<p>Another example with half-round cases only:</p>
<pre>
       1 digit        1 digit
Unrounded  "Standard" Rounding  "Gaussian rounding"
 27.25    27.3          27.2
 27.45    27.5          27.4
+ 27.55   + 27.6         + 27.6
=======   ======         ======
 82.25    82.4          82.2</pre>
<p>Again, the "Gaussian" rounding result is nearer from the
unrounded result than the "Standard" one.</p>
<p>René Rhéaume<br />
rener@moncourrier.com</p>
<hr width="95%" size="2" align="center" />
<p>Rd.</p></font>

