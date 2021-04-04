 
$URL = (Get-Content "autocapture_url.txt") -as [string[]]
 
$ie = New-Object -ComObject InternetExplorer.Application
$ie.Visible    = $true
$ie.AddressBar = $true

$dll_info = '[DllImport("user32.dll")] public static extern bool ShowWindowAsync(IntPtr hWnd, int nCmdShow);'
Add-Type -MemberDefinition $dll_info -Name NativeMethods -Namespace Win32
[Win32.NativeMethods]::ShowWindowAsync($ie.HWND, 3)  | Out-Null

Add-Type -AssemblyName Microsoft.VisualBasic
$window_process = Get-Process -Name "iexplore" | ? {$_.MainWindowHandle -eq $ie.HWND}
[Microsoft.VisualBasic.Interaction]::AppActivate($window_process.ID) | Out-Null
 
foreach($url in $URL){

    $ie.navigate($url); 
    while($ie.busy){
        Start-Sleep -milliseconds 100
    }
    Start-Sleep -Seconds 1
 
    Add-Type -AssemblyName System.Drawing
    $bmp = new-object -TypeName System.Drawing.Bitmap -ArgumentList $ie.Width, $ie.Height
    $graphics = [System.Drawing.Graphics]::FromImage($bmp)
    $graphics.CopyFromScreen($ie.Left, $ie.Top, 0, 0, $bmp.Size)
 
    $filename = Split-Path $url -Leaf
    $bmp.Save($filename + ".png")    
 
    $count = $count + 1

}

$ie.Quit();