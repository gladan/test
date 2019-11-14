Sub Automate_IE_Load_Page()
'This will load a webpage in IE
    Dim i As Long
    Dim URL As String
    Dim IE As Object
    Dim objElement As Object
    Dim objCollection As Object
 
    'Create InternetExplorer Object
    Set IE = CreateObject("InternetExplorer.Application")
 
    'Set IE.Visible = True to make IE visible, or False for IE to run in the background
    IE.Visible = True
 
    'Define URL
    URL = "https://moamelat.tbtb.ir"
 
    'Navigate to URL
    IE.Navigate URL
 
    ' Statusbar let's user know website is loading
    Application.StatusBar = URL & " is loading. Please wait..."
 
    ' Wait while IE loading...
    'IE ReadyState = 4 signifies the webpage has loaded (the first loop is set to avoid inadvertently skipping over the second loop)
    Do While IE.ReadyState = 4: DoEvents: Loop   'Do While
    Do Until IE.ReadyState = 4: DoEvents: Loop   'Do Until
 
    'Webpage Loaded
    Application.StatusBar = URL & " Loaded"
    
    'Unload IE
    Set IE = Nothing
    Set objElement = Nothing
    Set objCollection = Nothing
    
End Sub


'This Must go at the top of your module. It's used to set IE as the active window
Public Declare Function SetForegroundWindow Lib "user32" (ByVal HWND As Long) As Long
 
Sub Automate_IE_Enter_Data()
'This will load a webpage in IE
    Dim i As Long
    Dim URL As String
    Dim IE As Object
    Dim objElement As Object
    Dim objCollection As Object
    Dim HWNDSrc As Long
    
 
    'Create InternetExplorer Object
    Set IE = CreateObject("InternetExplorer.Application")
 
    'Set IE.Visible = True to make IE visible, or False for IE to run in the background
    IE.Visible = True
 
    'Define URL
    URL = "https://www.automateexcel.com/excel/vba"
 
    'Navigate to URL
    IE.Navigate URL
 
    ' Statusbar let's user know website is loading
    Application.StatusBar = URL & " is loading. Please wait..."
 
    ' Wait while IE loading...
    'IE ReadyState = 4 signifies the webpage has loaded (the first loop is set to avoid inadvertantly skipping over the second loop)
    Do While IE.ReadyState = 4: DoEvents: Loop
    Do Until IE.ReadyState = 4: DoEvents: Loop
 
    'Webpage Loaded
    Application.StatusBar = URL & " Loaded"
    
    'Get Window ID for IE so we can set it as activate window
    HWNDSrc = IE.HWND
    'Set IE as Active Window
    SetForegroundWindow HWNDSrc
    
    
    'Find & Fill Out Input Box
    n = 0
    
    For Each itm In IE.document.all
        If itm = "[object HTMLInputElement]" Then
        n = n + 1
            If n = 3 Then
                itm.Value = "orksheet"
                itm.Focus                             'Activates the Input box (makes the cursor appear)
                Application.SendKeys "{w}", True      'Simulates a 'W' keystroke. True tells VBA to wait
                                                      'until keystroke has finished before proceeding, allowing
                                                      'javascript on page to run and filter the table
                GoTo endmacro
            End If
        End If
    Next
    
    'Unload IE
endmacro:
    Set IE = Nothing
    Set objElement = Nothing
    Set objCollection = Nothing
    
End Sub

IE.document.getelementbyid("ID").value = "value"       'Find by ID
IE.document.getelementsbytagname("ID").value = "value"        'Find by tag
IE.document.getelementsbyclassname("ID").value = "value"      'Find by class
IE.document.getelementsbyname("ID").value = "value"           'Find by name

In the code above we use the event: Focus (itm.focus) to activate the cursor in the form.

You can find more examples of Object / Element Events, Methods, and Properties here: https://msdn.microsoft.com/en-us/library/ms535893(v=vs.85).aspx

Not all of these will work with every object / element and there may be quite a bit of trial and error when interacting with objects in IE.

