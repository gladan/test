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
