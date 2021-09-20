#region IMPORTS
[void][System.Reflection.Assembly]::LoadWithPartialName('presentationframework')
Add-Type -AssemblyName System.Windows.Forms
try {
    Set-ExecutionPolicy Bypass -Force -Confirm:$false
} catch {
    [System.Windows.Forms.MessageBox]::Show("Failed to set execution policy`n`n$($_.Exception.Message)",'Set Policy','OK','Error')
    Exit
}

#endregion

function Open-ExportForm {
    $syncHash = [hashtable]::Synchronized(@{})
    $newRunspace = [runspacefactory]::CreateRunspace()
    $syncHash.Runspace = $newRunspace
    $newRunspace.ApartmentState = "STA"
    $newRunspace.ThreadOptions = "ReuseThread"
    $newRunspace.Name = "expForm"
    $data = $newRunspace.Open()
    $newRunspace.SessionStateProxy.SetVariable("syncHash",$syncHash)
    $psCmd = [powershell]::Create().AddScript({
        [xml]$xaml = @"
        <Window
            xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
            xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
            Title="Azure Security Report Tool - Ready" Height="200" Width="600" MaxHeight="200" MaxWidth="600" MinHeight="200" MinWidth="600" WindowStartupLocation="CenterScreen">
            <Grid>
                <Label Content="Subscription ID:" HorizontalAlignment="Left" Margin="152,45,0,0" VerticalAlignment="Top"/>
                <Button Name="btnRun" Content="Run" HorizontalAlignment="Left" Margin="458,101,0,0" VerticalAlignment="Top" Width="100"/>
                <TextBox Name="txtSub" HorizontalAlignment="Left" Height="23" Margin="156,71,0,0" TextWrapping="Wrap" VerticalAlignment="Top" VerticalContentAlignment="Center" Width="402"/>
                <Image HorizontalAlignment="Left" Height="148" Margin="10,10,0,0" VerticalAlignment="Top" Width="132" Source="https://i.imgur.com/o1uQuq0.png" Stretch="Fill" StretchDirection="DownOnly"/>
            </Grid>
        </Window>
"@
        #region INITIALIZATION
        #---XAML parser---#
        $reader = (New-Object System.Xml.XmlNodeReader $xaml)
        $syncHash.Window = [Windows.Markup.XamlReader]::Load($reader)
        $xaml.SelectNodes("//*[@Name]") | % {$syncHash."$($_.Name)" = $syncHash.Window.FindName($_.Name)}
        #endregionW

        #region CONTROLS UPDATER
        $syncHash.Enable = $true
        $syncHash.Status
        $syncHash.StatusPh
        $updateControls = {
            if($syncHash.Status -ne $syncHash.StatusPh) {
                $syncHash.Window.Title = $syncHash.Status
                $syncHash.StatusPh = $syncHash.Status
            }

            $syncHash.btnRun.IsEnabled = $syncHash.Enable
            $syncHash.txtSub.IsEnabled = $syncHash.Enable
        }

        $syncHash.Window.Add_SourceInitialized({
            $timer = New-Object System.Windows.Threading.DispatcherTimer   
            $timer.Interval = [TimeSpan]"0:0:0.01"          
            $timer.Add_Tick($updateControls)            
            $timer.Start()                       
        })

        #endregion
        

        #region CONTROL EVENTS
        $syncHash.Window.Add_Closing({
            if($syncHash.stopwatch.IsRunning -eq $true) {
                $closeWindow = [System.Windows.Forms.MessageBox]::Show('The report is currently running. Closing the window will stop the report creation. Do you want to proceed?','Azure Security Report Tool','YesNo','Question')
                    if($closeWindow -eq "Yes") {
                        $rs = Get-Runspace | where {$_.Name -eq "btnRun"}
                        $rs.Close()
                        $rs.Dispose()
                        $_.Close()
                    } else {
                        $_.Cancel = $true
                    }
            }
        })

        $syncHash.btnRun.Add_Click({
            $syncHash.stopwatch = [System.Diagnostics.Stopwatch]::StartNew()
            $syncHash.SubscriptionId = $syncHash.txtSub.Text

            $btnRunspace = [runspacefactory]::CreateRunspace()
            $btnRunspace.ApartmentState = "STA"
            $btnRunspace.ThreadOptions = "ReuseThread"
            $btnRunspace.Name = "btnRun"
            $btnRunspace.Open()
            $btnRunspace.SessionStateProxy.SetVariable("syncHash",$syncHash)
            $cmdBtn = [powershell]::Create().AddScript({
                $syncHash.Enable = $false

                #region HELPER FUNCTIONS
                function Find-InTextFile {
                    [CmdletBinding(DefaultParameterSetName = 'NewFile')]
                    [OutputType()]
                    param (
                        [Parameter(Mandatory = $true)]
                        [ValidateScript({Test-Path -Path $_ -PathType 'Leaf'})]
                        [string[]]$FilePath,
                        [Parameter(Mandatory = $true)]
                        [string]$Find,
                        [Parameter()]
                        [string]$Replace,
                        [Parameter(ParameterSetName = 'NewFile')]
                        [ValidateScript({ Test-Path -Path ($_ | Split-Path -Parent) -PathType 'Container' })]
                        [string]$NewFilePath,
                        [Parameter(ParameterSetName = 'NewFile')]
                        [switch]$Force
                    )
                    begin {
                        $Find = [regex]::Escape($Find)
                    }
                    process {
                        try {
                            foreach ($File in $FilePath) {
                                if ($Replace) {
                                    if ($NewFilePath) {
                                        if ((Test-Path -Path $NewFilePath -PathType 'Leaf') -and $Force.IsPresent) {
                                            Remove-Item -Path $NewFilePath -Force
                                            (Get-Content $File) -replace $Find, $Replace | Add-Content -Path $NewFilePath -Force
                                        } elseif ((Test-Path -Path $NewFilePath -PathType 'Leaf') -and !$Force.IsPresent) {
                                            Write-Warning "The file at '$NewFilePath' already exists and the -Force param was not used"
                                        } else {
                                            (Get-Content $File) -replace $Find, $Replace | Add-Content -Path $NewFilePath -Force
                                        }
                                    } else {
                                        (Get-Content $File) -replace $Find, $Replace | Add-Content -Path "$File.tmp" -Force
                                        Remove-Item -Path $File
                                        Move-Item -Path "$File.tmp" -Destination $File
                                    }
                                } else {
                                    Select-String -Path $File -Pattern $Find
                                }
                            }
                        } catch {
                            Write-Error $_.Exception.Message
                        }
                    }
                }
                function Update-Status([string]$Stat) {
                    $baseTitle = "Azure Security Report Tool"
                    $syncHash.Status = "$baseTitle - $Stat"
                }
                #endregion

                if($syncHash.SubscriptionId -and ($syncHash.SubscriptionId -match '(?im)^[{(]?[0-9A-F]{8}[-]?(?:[0-9A-F]{4}[-]?){3}[0-9A-F]{12}[)}]?$')){
                    #---Install NuGet---#
                    Update-Status "Validating package provider..."
                    $syncHash.nuget = Get-PackageProvider | where Name -eq "Nuget"
                    if(!$syncHash.nuget){
                        try {
                            Update-Status "Installing package provider..."
                            $syncHash.nugetInstall = Install-PackageProvider -Name NuGet -RequiredVersion 2.8.5.201 -Force -ErrorAction Stop
                        } catch {
                            [System.Windows.Forms.MessageBox]::Show("Failed to install package provider`n`n$($_.Exception.Message)",'Install Module','OK','Error')
                            $syncHash.Enable = $true
                            Update-Status "Ready"
                            $syncHash.stopwatch.Stop()
                        }
                    }
                    
                    #---Install PowerShellGet---#
                    Update-Status "Validating module installer..."
                    $syncHash.psGet = Get-InstalledModule PowerShellGet
                    if(!$syncHash.psGet){
                        try {
                            Update-Status "Installing module installer..."
                            Install-Module -Name PowerShellGet -Force -ErrorAction Stop
                        } catch {
                            [System.Windows.Forms.MessageBox]::Show("Failed to install module psget`n`n$($_.Exception.Message)",'Install Module','OK','Error')
                            $syncHash.Enable = $true
                            Update-Status "Ready"
                            $syncHash.stopwatch.Stop()
                        }
                    }

                    #---Install PoshRS---#
                    Update-Status "Validating runspace manager..."
                    $syncHash.poshRs = Get-InstalledModule -Name PoshRSJob
                    if(!$syncHash.poshRs){
                        try {
                            Update-Status "Installing runspace manager..."
                            Install-Module -Name PoshRSJob -Force -Confirm:$false -ErrorAction Stop
                            $syncHash.poshRsInstall = Get-InstalledModule | where Name -eq "PoshRSJob"
                        } catch {
                            [System.Windows.Forms.MessageBox]::Show("Failed to install runspace module`n`n$($_.Exception.Message)",'Install Module','OK','Error')
                            $syncHash.Enable = $true
                            Update-Status "Ready"
                            $syncHash.stopwatch.Stop()
                        }
                    }

                    #---Install AzSK---#
                    Update-Status "Validating required modules..."
                    $syncHash.modules = Get-InstalledModule -Name AzSK
                    if(!$syncHash.modules) {
                        Update-Status "Installing required modules..."
                        Invoke-Expression 'cmd /c start powershell -WindowStyle hidden -Command { Install-Module -Name AzSK -AllowClobber -Force -Confirm:$false -SkipPublisherCheck }'
                    }

                    $i = 1
                    do {
                        Update-Status "Validating module - attempt $i..."
                        $azsk = Get-InstalledModule -Name AzSK
                        $i++
                    } while(!$azsk -and $i -le 3)

                    if(!$azsk) {
                        [System.Windows.Forms.MessageBox]::Show("Failed to install required modules",'Install Modules','OK','Error')
                        $syncHash.Enable = $true
                        Update-Status "Ready"
                        $syncHash.stopwatch.Stop()
                        Exit
                    } else {
                        #---Update policy agreement---#
                        Update-Status "Validating module settings..."
                        $modPath = (Get-Module -l azsk*).path | Split-Path
                        $policyPath = "$modPath\Framework\Abstracts\PrivacyNotice.ps1"
                        $polConf = '$input = "y"'
                        if(!(Find-InTextFile -FilePath $policyPath -Find $polConf)) {
                            try {
                                Update-Status "Updating module settings..."
                                Find-InTextFile -FilePath $policyPath -Find '$input = ""' -Replace $polConf
                            } catch {
                                [System.Windows.Forms.MessageBox]::Show("Failed to update policy agreement`n`n$($_.Exception.Message)",'Update Policy','OK','Error')
                                $syncHash.Enable = $true
                                Update-Status "Ready"
                                $syncHash.stopwatch.Stop()
                            }
                        }
                        
                        #---Generate reports---#
                        try {
                            Update-Status "Generating security report..."
                            $rsTime = (Get-Date).ToString("MMddHHmmss")
                            $syncHash.rsJob = Start-RSJob -Name "$($rsTime)RS" -ScriptBlock {
                                Param($subGuid)
                                Import-Module AzSK
                                Get-AzSKSubscriptionSecurityStatus -SubscriptionId $subGuid
                            } -ArgumentList $syncHash.SubscriptionId
                        } catch {
                            [System.Windows.Forms.MessageBox]::Show("Failed to create security report`n`n$($_.Exception.Message)",'Generate Report','OK','Error')
                            $syncHash.Enable = $true
                            Update-Status "Ready"
                            $syncHash.stopwatch.Stop()
                        }

                        #---Monitor job---#
                        do {} while($syncHash.rsJob.Completed -ne $true)

                        if($syncHash.rsJob.HasErrors){
                            [System.Windows.Forms.MessageBox]::Show("Failed to generate report -> $($syncHash.rsJob.Error[0].Exception.Message)",'Generate Report','OK','Error')
                        }
                        
                        Update-Status "Ready"
                        $syncHash.stopwatch.Stop()
                        $syncHash.Enable = $true
                    }
                } else {
                    [System.Windows.Forms.MessageBox]::Show("GUID invalid",'Error','OK','Error')
                    Update-Status "Ready"
                    $syncHash.stopwatch.Stop()
                    $syncHash.Enable = $true
                }
            })
            $cmdBtn.Runspace = $btnRunspace
            $cmdBtn.BeginInvoke() | Out-Null
        })
        #endregion
        
        $syncHash.Window.ShowDialog() | Out-Null
        $syncHash.Error += $Error
    })

    $psCmd.Runspace = $newRunspace
    $data = $psCmd.BeginInvoke()
    
    Return $syncHash
}

$prg = Open-ExportForm