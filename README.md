<div align="center">

## Fixing Popupmenu 'bug' from system tray


</div>

### Description

Simple code + explaination of why popup menus don't disapear like most app's menus when called from the system tray. Very hard problem, very simple answer. (this has no minimize to system tray code in it)
 
### More Info
 


<span>             |<span>
---                |---
**Submitted On**   |
**By**             |[SKoW](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByAuthor/skow.md)
**Level**          |Beginner
**User Rating**    |5.0 (149 globes from 30 users)
**Compatibility**  |VB 3\.0, VB 4\.0 \(16\-bit\), VB 4\.0 \(32\-bit\), VB 5\.0, VB 6\.0, VB Script, ASP \(Active Server Pages\) , VBA MS Access, VBA MS Excel
**Category**       |[Windows API Call/ Explanation](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByCategory/windows-api-call-explanation__1-39.md)
**World**          |[Visual Basic](https://github.com/Planet-Source-Code/PSCIndex/blob/master/ByWorld/visual-basic.md)
**Archive File**   |[](https://github.com/Planet-Source-Code/skow-fixing-popupmenu-bug-from-system-tray__1-21024/archive/master.zip)





### Source Code

<html>
<body>
<h2 align="center"><font color="#FF0000"><b>Fixing PopupMenu Fault when used in
a System Tray Project</b></font></h2>
<p>&nbsp;&nbsp;&nbsp; If you have ever sent your app to the system tray and created a click event to popup a menu using Popupmenu (menuname) you will be aware of the small 'bug'.</p>
<p>&nbsp;&nbsp;&nbsp; By default, your application <font color="#FF0000"><b> WILL NOT</b></font> be set as the active app so when you call the popupmenu function, it will create the menu out of focus. This means it cannot recieve the 'lost focus' message and will not close when you click somewhere else. This is a very simple thing to fix but nobody seems to use it. All you have to do is set the Form owner of the menu to Focus :)<br>
</p>
<p>eg: (from standard sys tray form_mousemove event)</p>
<hr>
<p><font color="#009933">'(place in module)</font><font color="#008080"><br>
</font><font color="#0000FF">Public</font> <font color="#0000FF"> Declare Function</font> SetForegroundWindow
<font color="#0000FF"> Lib</font> "user32" (<font color="#0000FF">ByVal</font> hwnd
<font color="#0000FF"> As</font> <font color="#0000FF">Long</font>) <font color="#0000FF"> As Long</font><br>
</p>
<hr>
<p><br>
<br>
<font color="#009933">'(place in form_mousemove event)</font><font color="#008080"><br>
</font><font color="#0000FF">Private Sub</font> Form_MouseMove(Button <font color="#0000FF"> As</font>
<font color="#0000FF">Integer</font>, Shift <font color="#0000FF"> As</font> <font color="#0000FF">Integer</font>, X
<font color="#0000FF"> As</font> <font color="#0000FF">Single</font>, Y <font color="#0000FF"> As
Single</font>)<br>
</p>
<p>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <font color="#0000FF">If</font> Me.WindowState
<font color="#0000FF"> =</font> vbMinimized then<br>
<font color="#009933">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
' window is minimized must be in system tray or MouseMove event would not
execute</font><br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <font color="#0000FF">Dim</font>
lngMsg <font color="#0000FF"> As Long</font><br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <font color="#0000FF">Dim</font> result
<font color="#0000FF"> As Long</font><br>
<font color="#009933">&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
' get the WM Message passed via X<br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
' since X is by default mes. in Twips,&nbsp;<br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
' devide it by the number of twips / pixel<br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
' so we recieve the proper value</font><br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; lngMsg <font color="#0000FF"> =</font> X / Screen.TwipsPerPixelX<br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;</p>
<p>&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <font color="#0000FF">Select Case</font> lngMsg<br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
<font color="#0000FF">case</font> WM_RBUTTONUP<font color="#009933"> ' right button&nbsp;</font><br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
SetForegroundWindow Me.hwnd<br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;
Popupmenu Me.mnuFile<br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <font color="#0000FF">end select</font><br>
&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<font color="#0000FF"> end if</font><br>
<font color="#0000FF">end sub</font></p>
<hr>
<p>&nbsp;</p>
<p> and it's that simple. Can't remember who told me this but thanks if it were ye. :)<br>
</p>
</body>
</html>

