## REPLACE START
$arrayOfUpdates=@(
    [pscustomobject]@{
        Type = 'net'
        FeatureVersion = '22H2'
        BuildNumber = 'na'
        PatchNumber = 'KB5031988'
        PatchLink = 'https://catalog.s.download.windowsupdate.com/d/msdownload/update/software/secu/2023/11/windows10.0-kb5031988-x64-ndp48_cf8f8281813f1474b04a83398f4615c08786c2db.msu'
    },[pscustomobject]@{
        Type = 'net'
        FeatureVersion = '21H2'
        BuildNumber = 'na'
        PatchNumber = 'KB5031988'
        PatchLink = 'https://catalog.s.download.windowsupdate.com/d/msdownload/update/software/secu/2023/11/windows10.0-kb5031988-x64-ndp48_cf8f8281813f1474b04a83398f4615c08786c2db.msu'
    },[pscustomobject]@{
        Type = 'cumulative'
        FeatureVersion = '22H2'
        BuildNumber = '19045.3693'
        PatchNumber = 'KB5032189'
        PatchLink = 'https://catalog.s.download.windowsupdate.com/d/msdownload/update/software/secu/2023/11/windows10.0-kb5032189-x64_0a3b690ba3fa6cd69a2b0f989f273cfeadba745f.msu'
    },[pscustomobject]@{
        Type = 'cumulative'
        FeatureVersion = '21H2'
        BuildNumber = '19044.3693'
        PatchNumber = 'KB5032189'
        PatchLink = 'https://catalog.s.download.windowsupdate.com/d/msdownload/update/software/secu/2023/11/windows10.0-kb5032189-x64_0a3b690ba3fa6cd69a2b0f989f273cfeadba745f.msu'
    }
)
## REPLACE END


$isSuccess = $false

function Install-KBUpdate {
    param (
        $patchNumber,
        $patchLink
    )
    $hotfix = Get-HotFix -Id $patchNumber -ErrorAction 0; 
    If ($hotfix) { 
        Write-Host "Patch $patchNumber is already instaled"
        return 0
    }
    Else {
        #download and install .Net patch
        $path = "C:\Temp\${patchNumber}.msu" ;
        Invoke-WebRequest -Uri $patchLink -OutFile $path ;
    
        if (Test-Path -Path $path -PathType Leaf) {
            Set-Location -Path "C:\Temp"
            New-Item -Path ".\" -Name "Update" -ItemType "directory";
            expand -F:* "C:\temp\${patchNumber}.msu" C:\Temp\Update\ ;
            $updateFile = Get-ChildItem -File "C:\Temp\Update\*$patchNumber*.cab" | Select-Object -ExpandProperty Name
            dism.exe /online /add-package /packagepath:"C:\Temp\Update\$updateFile" /quiet /norestart 
            remove-item "C:\Temp\${patchNumber}.msu" -Force;
            remove-item -LiteralPath C:\Temp\Update -Force -Recurse;
            Write-Host "Patch $patchNumber have been installed"
        }     

    }
    return 0
}

foreach ($patch in $arrayOfUpdates) {
    Install-KBUpdate -patchNumber $patch.PatchNumber -patchLink $patch.PatchLink
    if ($?) {
        $isSuccess = $true
    }
    else {
        $isSuccess = $false
    }
}



if ($isSuccess) {
    exit 0
}
else {
    exit 1
}
## REPLACE_DATE 2023-11-20 19:55:29
