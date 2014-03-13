
<#
    .Description
    Functionality:
    1. Connect to common shared paths: \\exstore, \\exfiles, \\products\public
    2. Install specified OUTLOOK
    Usage:
    1. Run the script in a powershell prompt window: PS D:\> .\InitTestEnvironment.ps1
    2. Input username and password.
    3. Select a certain version of OUTLOOK to install.
#>

###### Const ########
$OLK14SetupPath_64bit = "\\products\PUBLIC\PRODUCTS\Applications\User\Office_2010_SP1\64-bit\Office_ProfessionalPlus_2010_SP1\"
$OLK14SetupPath_32bit = "\\products\PUBLIC\PRODUCTS\Applications\User\Office_2010_SP1\32-bit\Office_ProfessionalPlus_2010_SP1\"
$OLK15SetupPath_32bit = "\\products\PUBLIC\PRODUCTS\Applications\User\Office_2013\English\MSI\32-Bit\Office_Professional_2013\"
###### Const ########

function NetUse($cred, [string] $path)
{
    if (($cred -eq $null) -or $path -eq $null)
    {
        throw 'Param $cred or $path is null.'
    }

    $nCred = $cred.GetNetWorkCredential()
    # Use $cred.UserName directly in the command will fail.(unknown username or bad password) What's the reason?
    $u = $cred.UserName

    net use $path $nCred.PassWord /USER:$u
    
    Write-Host "net use $path via user $u"
}

function NetUseCommonPaths($cred)
{
    $commonPaths = new-object System.Collections.Generic.List[string]
    $commonPaths.Add("\\exstore")
    $commonPaths.Add("\\exfiles")
    $commonPaths.Add("\\products")
    $commonPaths.Add("\\products\public")
	#$commonPaths.Add("\\atc-icebreaker\users\v-kail")

    Write-Host "Connecting to common shared paths..." -ForegroundColor Yellow
    foreach($p in $commonPaths)
    {
        NetUse $cred $p
    }
}

function GenOLKConfigXML()
{
    $OLKConfigFileFullPath = 'd:\OLKCustomConfig.xml'
    $xmlContent = @"
<Configuration>
  <Logging Type="standard" Path="%systemdrive%" Template="Office14(*).log" />
  <PIDKEY Value="" />
  <Display Level="Basic" CompletionNotice="Yes" AcceptEula="yes" />
  <OptionState Id="ACCESSFiles" State="absent" Children="force" />
  <OptionState Id="EXCELFiles" State="absent" Children="force" />
  <OptionState Id="GrooveFiles" State="absent" Children="force" />
  <OptionState Id="OneNoteFiles" State="absent" Children="force" />
  <OptionState Id="OUTLOOKFiles" State="Local" Children="force" />
  <OptionState Id="PPTFiles" State="absent" Children="force" />
  <OptionState Id="PubPrimary" State="absent" Children="force" />
  <OptionState Id="WORDFiles" State="absent" Children="force" />
  <OptionState Id="XDOCSFiles" State="absent" Children="force" />
  <OptionState Id="SHAREDFiles" State="absent" Children="force" />
  <OptionState Id="TOOLSFiles" State="absent" Children="force" />
</Configuration>
"@
    Write-Host "Generating OUTLOOK config file: $OLKconfigFileFullPath" -ForegroundColor Green
    new-item $OLKconfigFileFullPath -type file -Force -Value $xmlContent
    
    #return $OLK14ConfigFileFullPath
    #new-item returns the file full path?
}

function InstallOLK([string]$setupCmd)
{
    $OLKConfigFileFullPath = GenOLKConfigXML
    
    & $setupCmd /config $OLKConfigFileFullPath

    Write-Host "Installing OUTLOOK..." -ForegroundColor Green
}

function main()
{
    Write-Host 'Please input your credential' -ForegroundColor Yellow
    $cred = Get-Credential

    NetUseCommonPaths($cred)

    [boolean]$reTry = $true
    [string]$OLKSetupPath = ""
    while ($reTry)
    {
        Write-Host "Please input the Numbder (1, 2 or 3) to install according version of OUTLOOK:" -ForegroundColor Yellow
        Write-Host "======================================================================"
        Write-Host "  1(64bitOUTLOOK 2010) | 2(32bitOUTLOOK 2010) | 3(32bitOUTLOOK 2013)  "
        Write-Host "======================================================================"
        $olkVersion = Read-Host
        switch ($olkVersion)
        {
            "1"
            {
                $OLKSetupPath = $OLK14SetupPath_64bit
                $reTry = $false
                break
            }
            "2"
            {
                $OLKSetupPath = $OLK14SetupPath_32bit
                $reTry = $false
                break
            }
            "3"
            {
                $OLKSetupPath = $OLK15SetupPath_32bit
                $reTry = $false
                break
            }
            default
            {
                Write-Host "Wrong input. Please retry." -ForegroundColor Red
                $reTry = $true
                break
            }
        }
    }
    
    [boolean]$pathAccessable = Test-Path $OLKSetupPath
    if (!$pathAccessable)
    {
        NetUse($cred, $OLKSetupPath)
    }


    if ($pathAccessable -or (Test-Path $OLKSetupPath))
    {
        $OLKSetupCmd = $OLKSetupPath + "setup.exe"
        InstallOLK($OLKSetupCmd)
    }
    else
    {
        Write-Host "$OLKSetupPath is not accessable." -ForegroundColor Red
    }
}

main