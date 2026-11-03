# Installing Excel-100x-Pro.xlam (desktop Excel, Windows)

## For one computer
1. Copy **Excel-100x-Pro.xlam** to a folder that will not be deleted, for example
   `C:\Users\<you>\AppData\Roaming\Microsoft\AddIns\`
2. Open Excel. Go to **File → Options → Add-ins**.
3. At the bottom, next to *Manage*, pick **Excel Add-ins** and click **Go...**
4. Click **Browse...**, select the file, and click **OK**. Make sure **Excel 100x Pro** is ticked.
5. A new ribbon tab, **100x Pro Tools**, appears with every tool in it.

If the ribbon does not appear, close Excel completely and reopen it. If Windows blocked the
downloaded file: right-click it → **Properties** → tick **Unblock** → **OK**.

## For a whole company
Put the file on a read-only network share, then either:
- tell staff to follow the steps above pointing at the share, or
- deploy it with Group Policy / Intune by copying it to each PC's `%APPDATA%\Microsoft\AddIns`
  and adding a registry value under
  `HKCU\Software\Microsoft\Office\16.0\Excel\Options` named `OPEN` (or `OPEN1`, `OPEN2`, ...)
  with the value `/R "C:\path\Excel-100x-Pro.xlam"`.

Also add the folder as a **Trusted Location** (File → Options → Trust Center → Trust Center
Settings → Trusted Locations) so no macro warning appears.

## Mac or Excel on the web
The `.xlam` is Windows-only. Use the web edition instead: <https://excel-100x-pro.vercel.app/>
(see README.md). The two can be installed side by side.

## Updating
Replace the file with a newer one while Excel is closed. Settings and your data are untouched.

## Removing
File → Options → Add-ins → Manage: Excel Add-ins → Go... → untick **Excel 100x Pro**.
