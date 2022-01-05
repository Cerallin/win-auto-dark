' Copyright (c) 2022 Cerallin <cerallin@cerallin.top>
'
' Permission is hereby granted, free of charge, to any person obtaining a copy
' of this software and associated documentation files (the "Software"), to deal
' in the Software without restriction, including without limitation the rights
' to use, copy, modify, merge, publish, distribute, sublicense, and/or sell
' copies of the Software, and to permit persons to whom the Software is
' furnished to do so, subject to the following conditions:
'
' The above copyright notice and this permission notice shall be included in all
' copies or substantial portions of the Software.
'
' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
' IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
' FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
' AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
' LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM,
' OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE
' SOFTWARE.

Option Explicit ' Strict mode.
' For CJK users: Please ensure that the encoding of this file is GBK.
' Please check the EOL Sequence: CRLF ("\r\n").
' VBS is not case sensitive, so all the objects, methods and attributes are camel-cased,
' and all the other variables are under-score-cased.

Dim light_off_time, light_on_time
Dim reg, target_theme

' Turn on dark mode after 21:00 (21 * 60 + 0 = 1260),
' and turn off before 6:00 (6 * 60 + 0 = 360).
light_off_time = 1260
light_on_time = 360

Class MyTime
    Private h
    Private m

    ' Initialization subs donot have any parameters.
    Private Sub Class_Initialize()
        ' Get current time
        Dim now_time
        now_time = Now

        '
        h = Hour(now_time)
        m = Minute(now_time)
    End Sub

    ' Get mintes of the day
    Public Property Get time()
        ' VBS returns with the form `function_name = expression`.
        time = h * 60 + m
    End Property
End Class

Class MyReg
    Private reg
    Private regPath

    ' Initialize shell for registry read/write.
    Private Sub Class_Initialize()
        Set RegObj = WScript.CreateObject("WScript.Shell")
        regPath = "HKCU\SOFTWARE\Microsoft\Windows\CurrentVersion\Themes\Personalize\"
    End Sub

    ' Make reg assignable.
    Private Property Set RegObj(obj)
        Set reg = obj
    End Property

    ' Get value of a variable from Windows registry.
    Private Function GetVar(var_name)
        GetVar = reg.RegRead(regPath & var_name)
    End Function

    ' Set value of a variable of Windows registry.
    Private Function SetVar(var_name, value)
        ' It is important to specify "DWORD" as the variable's type
        SetVar = reg.RegWrite(regPath & var_name, value, "REG_DWORD")
    End Function

    ' Get value of "AppsUseLightTheme".
    Public Property Get AppTheme()
        AppTheme = GetVar("AppsUseLightTheme")
    End Property

    ' Set value of "AppsUseLightTheme",
    ' using "Let" instead of "Set" means it's not an object.
    Public Property Let AppTheme(theme)
        AppTheme = SetVar("AppsUseLightTheme", theme)
    End Property

    ' Get value of "SystemUsesLightTheme".
    Public Property Get SysTheme()
        SysTheme = GetVar("SystemUsesLightTheme")
    End Property

    ' Set value of "SystemUsesLightTheme".
    Public Property Let SysTheme(theme)
        SysTheme = SetVar("SystemUsesLightTheme", theme)
    End Property

End Class

Function get_target_theme(cur_time)
    Dim tval
    tval = cur_time.time
    ' Calculate if current time is in the day time
    ' and convert the result into integer
    get_target_theme = CInt(tval < light_off_time AND tval > light_on_time)
End Function

' Start processing

' Target theme: to be light or dark?
target_theme = get_target_theme(new MyTime)
' Instantiate MyReg
Set reg = New MyReg

If reg.AppTheme <> target_theme Then
    reg.AppTheme = target_theme
End If

If reg.SysTheme <> target_theme Then
    reg.SysTheme = target_theme
End If
