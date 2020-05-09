Remove-Item ..\QuickLook.Plugin.PostScriptViewer.qlplugin -ErrorAction SilentlyContinue

$files = Get-ChildItem -Path ..\bin\Release\ -Exclude *.pdb,*.xml
Compress-Archive $files ..\QuickLook.Plugin.PostScriptViewer.zip
Move-Item ..\QuickLook.Plugin.PostScriptViewer.zip ..\QuickLook.Plugin.PostScriptViewer.qlplugin