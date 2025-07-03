Here’s how you can convert the key elements from your VB `.xaml.vb` code-behind into a PowerShell GUI script using **WPF/XAML with PSGUI**, which allows embedding PowerShell logic behind XAML-defined interfaces:

---

## 1. Load and parse the XAML

Assume your XAML defines a window named `MainWindow`, includes a `DataGrid` named `membersDataGrid`, and has a draggable `Border`:

```powershell
# Load XAML
Add-Type -AssemblyName PresentationFramework
[xml]$xaml = Get-Content 'MainWindow.xaml'

$reader   = (New-Object System.Xml.XmlNodeReader $xaml)
$window   = [Windows.Markup.XamlReader]::Load($reader)
```

---

## 2. Hook up Code‑Behind functionality

Attach PowerShell event handlers similar to VB’s:

```powershell
# Find named controls
$grid = $window.FindName('membersDataGrid')
$border = $window.FindName('DragBorder')  # Suppose the draggable border has x:Name="DragBorder"

# Enable dragging the window
$border.Add_MouseDown({
    if ($_.ChangedButton -eq 'Left') { $window.DragMove() }
})

# Double-click to toggle maximize/restore
$window.IsMaximized = $false
$border.Add_MouseLeftButtonDown({
    if ($_.ClickCount -eq 2) {
        if ($window.IsMaximized) {
            $window.WindowState = 'Normal'; $window.Width = 1080; $window.Height = 720
            $window.IsMaximized = $false
        } else {
            $window.WindowState = 'Maximized'
            $window.IsMaximized = $true
        }
    }
})
```

---

## 3. Populate the DataGrid

Just like the VB version uses an `ObservableCollection`, in PowerShell you can build a `[System.Collections.ObjectModel.ObservableCollection[PSCustomObject]]`:

```powershell
$members = New-Object System.Collections.ObjectModel.ObservableCollection[PSCustomObject]

# Helper to convert hex color string to Brush
function To-Brush($hex) {
    $conv = New-Object System.Windows.Media.BrushConverter
    return $conv.ConvertFromString($hex)
}

# Sample data
$members.Add([PSCustomObject]@{Number='1'; Character='Q'; BgColor=To-Brush '#ff6551'; Name='wetr'; Position='ert'; Email='123'; Phone='123'})
# ...repeat for other members...

# Assign to DataGrid
$grid.ItemsSource = $members
```

Make sure your XAML binds DataGrid columns to properties like `Number`, `Character`, `Name`, and uses `BgColor` for styling rows or cells.

---

## 4. Show the Window

```powershell
# Center and show
$window.WindowStartupLocation = 'CenterScreen'
$window.ShowDialog() | Out-Null
```

---

### ✅ Complete Example

```powershell
Add-Type -AssemblyName PresentationFramework

[xml]$xaml = Get-Content 'MainWindow.xaml'
$reader = New-Object System.Xml.XmlNodeReader $xaml
$window = [Windows.Markup.XamlReader]::Load($reader)

# Locate controls
$grid = $window.FindName('membersDataGrid')
$border = $window.FindName('DragBorder')

# Drag and double-click maximize logic
$border.Add_MouseDown({ if ($_.ChangedButton -eq 'Left') { $window.DragMove() } })
$window.IsMaximized = $false
$border.Add_MouseLeftButtonDown({
    if ($_.ClickCount -eq 2) {
        if ($window.IsMaximized) {
            $window.WindowState='Normal'; $window.Width=1080; $window.Height=720; $window.IsMaximized=$false
        } else {
            $window.WindowState='Maximized'; $window.IsMaximized=$true
        }
    }
})

# Populate DataGrid
function To-Brush($hex){
    $conv = New-Object System.Windows.Media.BrushConverter
    $conv.ConvertFromString($hex)
}
$members = New-Object System.Collections.ObjectModel.ObservableCollection[PSCustomObject]
# Add items (repeat as needed)
$members.Add([PSCustomObject]@{Number='1'; Character='Q'; BgColor=To-Brush '#ff6551'; Name='wetr'; Position='ert'; Email='123'; Phone='123'})
# ...
$grid.ItemsSource = $members

# Display window
$window.WindowStartupLocation='CenterScreen'
$window.ShowDialog() | Out-Null
```

---

### 🔧 Tips & Next Steps

| Suggestion                                                                                                            | Why It Helps                                                                                                                                                                                                                             |
| --------------------------------------------------------------------------------------------------------------------- | ---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------- |
| **Install PSGUI module**: Run `Install-Module PSGUI; Import-Module PSGUI` to simplify loading XAML and wiring events. | It provides helpers like `Initialize-XAMLDialog` and auto-generates PowerShell variables for named controls. ([subscription.packtpub.com][1], [blog.csdn.net][2], [powerintheshell.com][3], [ideone.com][4], [powershellgallery.com][5]) |
| **Enhance styling/data-binding**: Use XAML data templates and styles, binding cells to `BgColor` for coloring rows.   | Keeps separation between layout (XAML) and logic (PowerShell).                                                                                                                                                                           |
| **Focus on reusable patterns**: Abstract repetitive event registration or object creation into functions/modules.     | Simplifies future maintenance and improves readability.                                                                                                                                                                                  |

---

### ✅ Next Steps for You

1. **Provide your actual `MainWindow.xaml`**, so I can align the PowerShell script with your named controls and layout.
2. **Confirm which events or features** (like dragging, double-click, grid loading) you want fully converted.
3. I’ll help complete the script based on your real XAML, ensuring it matches exactly what you’ve built in the `.vb` version.

Just paste that XAML, and I’ll finalize the PowerShell port for you!

[1]: https://subscription.packtpub.com/book/programming/9781788838986/17/ch17lvl1sec172/creating-xaml-guis-with-psgui?utm_source=chatgpt.com "Learn PowerShell Core 6.0"
[2]: https://blog.csdn.net/weixin_48239221/article/details/130907355?utm_source=chatgpt.com "WPF界面设计-CSDN博客"
[3]: https://powerintheshell.com/2016/07/22/psgui-hello-world-howto/?utm_source=chatgpt.com "PSGUI – Hello World! HowTo | Power in the shell"
[4]: https://ideone.com/zKXvF?utm_source=chatgpt.com "zKXvF - Online C# Compiler & Debugging Tool - Ideone.com"
[5]: https://www.powershellgallery.com/packages/PSGUI/0.21/Content/PSGUI_Manager/PSGUI_Manager.xaml?utm_source=chatgpt.com "PowerShell Gallery | PSGUI_Manager/PSGUI_Manager.xaml 0.21"
