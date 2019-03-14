$ScriptVersion = "1.0.0"
#region GUI
#ERASE ALL THIS AND PUT XAML BELOW between the @" "@
$inputXML = @"
<Window x:Class="MSIX_Commander.MainWindow"
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        xmlns:local="clr-namespace:MSIX_Toolkit"
        mc:Ignorable="d"
        Title="MSIX Commander" Width="800" Height="600">
    <Grid>
        <Grid.ColumnDefinitions>
            <ColumnDefinition/>
        </Grid.ColumnDefinitions>
        <Label x:Name="Label_Messages" Content="Messages:" HorizontalAlignment="Left" Height="27" Margin="0,8,0,0" VerticalAlignment="Top" Width="64"/>
        <TextBox x:Name="TextBox_Messages" Height="23" Margin="69,10,10,0" TextWrapping="Wrap" VerticalAlignment="Top" IsEnabled="False"/>
        <TabControl Margin="0,40,10,29.5">
            <TabItem Header="Installed">
                <Grid Background="#FFE5E5E5">
                    <Grid.RowDefinitions>
                        <RowDefinition/>
                    </Grid.RowDefinitions>
                    <ListView x:Name="ListView_MSIXPackages" Margin="0,87,0,0" VerticalAlignment="Top" HorizontalAlignment="Left" MinWidth="352">
                        <ListView.View>
                            <GridView>
                                <GridViewColumn Header="Name" DisplayMemberBinding ="{Binding Name}"/>
                                <GridViewColumn Header="Version" DisplayMemberBinding ="{Binding Version}"/>
                                <GridViewColumn Header="Publisher" DisplayMemberBinding ="{Binding Publisher}"/>
                                <GridViewColumn Header="InstallLocation" DisplayMemberBinding ="{Binding InstallLocation}"/>
                            </GridView>
                        </ListView.View>
                    </ListView>
                    <Button x:Name="Button_GetSoftware" Content="Get Software" HorizontalAlignment="Left" Margin="0,32,0,0" VerticalAlignment="Top" Width="93"/>
                    <Button x:Name="Button_Filter" Content="Filter" HorizontalAlignment="Left" Margin="560,29,0,0" VerticalAlignment="Top" Width="75"/>
                    <TextBox x:Name="TexBox_Filter" HorizontalAlignment="Left" Height="23" Margin="306,29,0,0" TextWrapping="Wrap" Text="" VerticalAlignment="Top" Width="249"/>
                    <Label x:Name="Label_Filter" Content="Filter:" HorizontalAlignment="Left" Height="23" VerticalAlignment="Top" Width="46" Margin="270,29,0,0"/>
                    <CheckBox x:Name="CheckBox_EnterpriseSigned" Content="Only Enterprise signed" HorizontalAlignment="Left" Margin="112,34,0,0" VerticalAlignment="Top" IsChecked="True"/>
                    <Button x:Name="Button_OpenInstallLocation" Content="Open Installlocation" HorizontalAlignment="Left" VerticalAlignment="Top" Width="114" Margin="0,58,0,0"/>
                    <Button x:Name="Button_OpenUserData" Content="Open Userdata" HorizontalAlignment="Left" Margin="119,58,0,0" Width="114" Height="20" VerticalAlignment="Top"/>
                    <Label x:Name="Label_Installed" Content="Work with installed MSIX Packages" HorizontalAlignment="Left" FontWeight="Bold" Height="26" VerticalAlignment="Top" Margin="-2,0,0,0"/>
                    <Button x:Name="Button_Uninstall" Content="Uninstall" HorizontalAlignment="Left" Margin="238,58,0,0" VerticalAlignment="Top" Width="114"/>
                </Grid>
            </TabItem>
            <TabItem x:Name="Tab_certificate" Header="Certificate">
                <Grid Background="#FFE5E5E5" Margin="0,0,0,-3">
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition/>
                    </Grid.ColumnDefinitions>
                    <Button x:Name="Button_CreatCert" Content="Creat selfsigned cert" Margin="326,173,0,0" VerticalAlignment="Top" Height="33" Width="128" HorizontalAlignment="Left"/>
                    <Label x:Name="Label_CreatCert" Content="Creat a selfsigned cert" HorizontalAlignment="Left" VerticalAlignment="Top" FontWeight="Bold" Margin="-2,0,0,0"/>
                    <TextBlock x:Name="TexBlock_CreatCert" HorizontalAlignment="Left" Margin="3,26,0,0" TextWrapping="Wrap" Text="This will creat a pfx file, with that and the password you can sign your MSIX files and a cer file that you need to install on your test machine." VerticalAlignment="Top" Height="27" Width="758"/>
                    <Button x:Name="Button_SelectCertStoreFolder" Content="Browse" Margin="0,132,10,0" VerticalAlignment="Top" Height="26" HorizontalAlignment="Right" Width="49"/>
                    <Label x:Name="Label_CreatCertPublisherName" Content="Publishername:" HorizontalAlignment="Left" Margin="0,60,0,0" VerticalAlignment="Top"/>
                    <TextBox x:Name="TextBox_NewCertPublisherName" HorizontalAlignment="Left" Height="26" Margin="96,60,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="256"/>
                    <Label x:Name="Label_CreatCertFriendlyName" Content="Friendlyname:" HorizontalAlignment="Left" Margin="368,60,0,0" VerticalAlignment="Top"/>
                    <TextBox x:Name="TextBox_NewCertFriendlyName" HorizontalAlignment="Left" Height="26" Margin="464,60,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="254"/>
                    <Label x:Name="Label_CreatCertPassword" Content="Password:" HorizontalAlignment="Left" Margin="0,91,0,0" VerticalAlignment="Top"/>
                    <PasswordBox x:Name="PasswordBox_NewCertPassword" HorizontalAlignment="Left" Margin="96,91,0,0" VerticalAlignment="Top" Width="256" Height="26"/>
                    <Label x:Name="Label_CreatCertFileName" Content="Filename:" HorizontalAlignment="Left" Margin="368,91,0,0" VerticalAlignment="Top"/>
                    <TextBox x:Name="TextBox_NewCertFileName" HorizontalAlignment="Left" Height="26" Margin="464,91,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Width="254"/>
                    <Label x:Name="Label_CreatCertFolderName" Content="Export Folder:" HorizontalAlignment="Left" Margin="0,132,0,0" VerticalAlignment="Top"/>
                    <TextBox x:Name="TextBox_NewCertFolderName" Height="26" Margin="96,132,64,0" TextWrapping="Wrap" VerticalAlignment="Top"/>
                    <Label x:Name="Label_ImportCert" Content="Import a selfsigned cert*" HorizontalAlignment="Left" VerticalAlignment="Top" FontWeight="Bold" Margin="-2,231,0,0"/>
                    <Button x:Name="Button_InstallCert" Content="Install cert*" Margin="326,333,0,0" VerticalAlignment="Top" Height="33" Width="128" HorizontalAlignment="Left"/>
                    <Button x:Name="Button_SelectCertToImport" Content="Browse" Margin="0,292,7,0" VerticalAlignment="Top" Height="26" HorizontalAlignment="Right" Width="48"/>
                    <Label x:Name="Label_SelectedCert" Content="Selected Cert:" HorizontalAlignment="Left" Margin="3,292,0,0" VerticalAlignment="Top"/>
                    <TextBox x:Name="TextBox_SelectedCert" Height="26" Margin="100,292,61,0" TextWrapping="Wrap" VerticalAlignment="Top"/>
                    <TextBlock x:Name="TexBlock_InstallCert" HorizontalAlignment="Left" Margin="3,257,0,0" TextWrapping="Wrap" Text="This will install the cert to the TrustedPeople Store of the local machine. This requires Administrator rights." VerticalAlignment="Top" Height="27" Width="758"/>
                </Grid>
            </TabItem>
            <TabItem x:Name="Tab_Testing" Header="Testing">
                <Grid Background="#FFE5E5E5" Margin="0,0,0,-183">
                    <Button x:Name="Button_TestSingleMSIX" Content="Test MSIX" Margin="325,209,0,0" VerticalAlignment="Top" Height="33" Width="128" HorizontalAlignment="Left"/>
                    <Label x:Name="Label_TestingSingle" Content="Test a single MSIX" HorizontalAlignment="Left" VerticalAlignment="Top" FontWeight="Bold" Margin="0,102,0,0"/>
                    <TextBlock x:Name="TexBlock_TestSingle" HorizontalAlignment="Left" Margin="6,128,0,0" TextWrapping="Wrap" Text="This will let you select one single MSIX File, it will then install the MSIX and tell you what shortcuts are new and let you open those to test your application. Once your done it will uninstall the MSIX again." VerticalAlignment="Top" Height="36" Width="758"/>
                    <Button x:Name="Button_TestSelectSingleMSIXFile" Content="Browse" Margin="0,169,10,0" VerticalAlignment="Top" Height="26" HorizontalAlignment="Right" Width="49"/>
                    <Label x:Name="Label_TestSingleFile" Content="Selected MSIX:" HorizontalAlignment="Left" Margin="0,169,0,0" VerticalAlignment="Top"/>
                    <TextBox x:Name="TextBox_TestSelectedSingleMSIX" Height="20" Margin="91,173,69,0" TextWrapping="Wrap" VerticalAlignment="Top" MaxHeight="20"/>
                    <Button x:Name="Button_TestMultipleMSIX" Content="Test multiple MSIX files" Margin="320,370,0,0" VerticalAlignment="Top" Height="33" Width="140" HorizontalAlignment="Left"/>
                    <Label x:Name="Label_TestingMultiple" Content="Test multiple MSIX files" HorizontalAlignment="Left" VerticalAlignment="Top" FontWeight="Bold" Margin="0,251,0,0"/>
                    <TextBlock x:Name="TexBlock_TestMultiple" HorizontalAlignment="Left" Margin="6,277,0,0" TextWrapping="Wrap" Text="This will let you select a folder, the tool will then search for all MSIX files in this folder and subfolders. It will then let you select witch files you want to install and then install those MSIX files one after the other and tell you what shortcuts are new so that you can test your application. Once you're done with that application, it will uninstall that MSIX again and install the next one." VerticalAlignment="Top" Height="54" Width="758"/>
                    <Button x:Name="Button_TestSelectMultipleFolder" Content="Browse" HorizontalAlignment="Right" Margin="0,332,10,0" VerticalAlignment="Top" Width="49" Height="26"/>
                    <Label x:Name="Label_TestFolderForMultiple" Content="Selected folder:" HorizontalAlignment="Left" Margin="0,332,0,0" VerticalAlignment="Top"/>
                    <TextBox x:Name="TextBox_TestSelectedMultipleMSIXFolder" Height="20" Margin="91,336,69,0" TextWrapping="Wrap" VerticalAlignment="Top" MaxHeight="20"/>
                    <Button x:Name="Button_TestOnlyInstall" Content="Install MSIX" Margin="325,70,0,0" VerticalAlignment="Top" Height="33" Width="128" HorizontalAlignment="Left"/>
                    <Label x:Name="Label_TestingInstallMSIX" Content="Only install MSIX" HorizontalAlignment="Left" VerticalAlignment="Top" FontWeight="Bold" Margin="0,-5,0,0"/>
                    <TextBlock x:Name="TexBlock_TestInstallMSIX" HorizontalAlignment="Left" Margin="6,21,0,0" TextWrapping="Wrap" Text="This will let you select one single MSIX file and install it. If you want to uninstall it again go to the tab Installed." VerticalAlignment="Top" Height="20" Width="758"/>
                    <Button x:Name="Button_TestSelectInstallMSIXFile" Content="Browse" Margin="0,37,10,0" VerticalAlignment="Top" Height="26" HorizontalAlignment="Right" Width="49"/>
                    <Label x:Name="Label_TestInstallMSIX" Content="Selected MSIX:" HorizontalAlignment="Left" Margin="0,37,0,0" VerticalAlignment="Top"/>
                    <TextBox x:Name="TextBox_TestSelectedInstallMSIX" Height="20" Margin="91,41,69,0" TextWrapping="Wrap" VerticalAlignment="Top" MaxHeight="20"/>
                </Grid>
            </TabItem>
            <TabItem x:Name="Tab_PackagingMachien" Header="PackagingMachien">
                <Grid Background="#FFE5E5E5" Margin="0,0,0,-183">
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition/>
                    </Grid.ColumnDefinitions>
                    <Button x:Name="Button_PackagingToolDriverCheckStatus" Content="Check state*" Margin="141,74,0,0" VerticalAlignment="Top" Height="33" Width="128" HorizontalAlignment="Left"/>
                    <Label x:Name="Label_PackagingToolDriverCheck" Content="Check Packaging Tool Driver state" HorizontalAlignment="Left" VerticalAlignment="Top" FontWeight="Bold" Margin="0,-5,0,0"/>
                    <TextBlock x:Name="TexBlock_PackagingToolDriverCheck" HorizontalAlignment="Left" Margin="6,21,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Height="36" Width="758" Text="This will let you check the state of the Packaging Tool Driver FeatureOnDemand (FOD). You can also install and uninstall it. Installing will take some time, please be patient. This requires Administrator rights. It only works with Windows 10 1809 or higher."/>
                    <Button x:Name="Button_PackagingToolDriverInstall" Content="Install from WindowsUpdate*" Margin="307,74,0,0" VerticalAlignment="Top" Height="33" HorizontalAlignment="Left" Width="164"/>
                    <Button x:Name="Button_PackagingToolDriverUninstall" Content="Uninstall*" Margin="506,74,0,0" VerticalAlignment="Top" Height="33" Width="128" HorizontalAlignment="Left"/>
                </Grid>
            </TabItem>
        </TabControl>
        <Label x:Name="Label_AdminRights" Content="*Those actions require the tool to be executed as an administrator" Margin="0,0,375,-1.5" Height="26" VerticalAlignment="Bottom"/>
        <Label x:Name="Label_Version" Content="Version:" HorizontalAlignment="Right" Margin="0,0,0,0.5" RenderTransformOrigin="2.747,0.588" Height="24" VerticalAlignment="Bottom"/>
    </Grid>
</Window>
"@    
 
$inputXML = $inputXML -replace 'mc:Ignorable="d"','' -replace "x:N",'N'  -replace '^<Win.*', '<Window'
 
[void][System.Reflection.Assembly]::LoadWithPartialName('presentationframework')
[xml]$XAML = $inputXML
#Read XAML
 
$reader=(New-Object System.Xml.XmlNodeReader $xaml)
try{$Form=[Windows.Markup.XamlReader]::Load( $reader )}
catch{Write-Host "Unable to load Windows.Markup.XamlReader. Double-check syntax and ensure .net is installed."}
 
#===========================================================================
# Load XAML Objects In PowerShell
#===========================================================================
 
$xaml.SelectNodes("//*[@Name]") | %{Set-Variable -Name "WPF$($_.Name)" -Value $Form.FindName($_.Name)}
 
Function Get-FormVariables{
<#
#Remove Comment to Debug

if ($global:ReadmeDisplay -ne $true){Write-host "If you need to reference this display again, run Get-FormVariables" -ForegroundColor Yellow;$global:ReadmeDisplay=$true}

write-host "Found the following interactable elements from our form" -ForegroundColor Cyan
get-variable WPF*

#>

}
 
Get-FormVariables

#endregion

#region Initialize
$ThisScriptParentPath = $MyInvocation.MyCommand.Path -replace $myInvocation.MyCommand.Name,""
$ThisScriptName = $myInvocation.MyCommand.Name

#Set Version
$WPFLabel_Version.Content = "Version: $ScriptVersion"

#Clear some Stuff First
$MSIXData =$null

#Check if the OS Version is supported
$OSVersion = (Get-WmiObject -class Win32_OperatingSystem ).Version

[int]$Major = $OSVersion.split(".")[0]
[int]$Build = $OSVersion.split(".")[2]

If($Major -ge 10){
    If($Build -ge 17763){
        $WPFTextBox_Messages.Text ="Youre Windows Version supports all features of this tool."
        $WPFTextBox_Messages.Foreground = "Green"
    }
    ElseIf($Build -eq 17134){
            $WPFTextBox_Messages.Text ="You are running Windows 10 1803. In this version not all features of this tool are working."
            $WPFTextBox_Messages.Foreground = "DarkOrange"
    }
    ElseIf($Build -eq 16299){
            $WPFTextBox_Messages.Text ="You are running Windows 10 1709. In this version not all features of this tool are working."
            $WPFTextBox_Messages.Foreground = "DarkOrange"
    }
    else{
            $WPFTextBox_Messages.Text ="You are running Windows $OSVersion. This tool is not working with your Version of Windows. You need W10 1709 or higer."
            $WPFTextBox_Messages.Foreground = "Red"
    }
}
else{
        $WPFTextBox_Messages.Text ="You are running Windows $OSVersion. This tool is not working with your Version of Windows. You need W10 1709 or higer."
        $WPFTextBox_Messages.Foreground = "Red"
}




#endregion

#region Functions


Function Install-PackagingToolDriver{

    try{
        $PackagingToolDriver = Get-WindowsCapability -Online | where  Name -Like "Msix.PackagingTool.Driver*"


        $Name = $PackagingToolDriver.Name
        $State = $PackagingToolDriver.State

        If($State -eq "NotPresent"){

            Add-WindowsCapability -Online -Name $Name -ErrorAction Stop

            $WPFTextBox_Messages.Text =  ("Installed $Name")
            $WPFTextBox_Messages.Foreground = "Black"
        }
        else{
            $WPFTextBox_Messages.Text =  ("The $Name state is already $State")
            $WPFTextBox_Messages.Foreground = "DarkOrange"
        }
    }
    catch{
        $ErrorMessage = $_.Exception.Message
        $WPFTextBox_Messages.Text =  ("Couldn't install the PackagingToolDriver / $ErrorMessage")
        $WPFTextBox_Messages.Foreground = "Red"
    }
}

Function Uninstall-PackagingToolDriver{

    try{
        $PackagingToolDriver = Get-WindowsCapability -Online | where  Name -Like "Msix.PackagingTool.Driver*"


        $Name = $PackagingToolDriver.Name
        $State = $PackagingToolDriver.State

        If($State -eq "Installed"){
            Remove-WindowsCapability -Online -Name $Name -ErrorAction Stop

            $WPFTextBox_Messages.Text =  ("Uninstalled $Name")
            $WPFTextBox_Messages.Foreground = "Black"
        }
        else{
            $WPFTextBox_Messages.Text =  ("The $Name state is already $State")
            $WPFTextBox_Messages.Foreground = "DarkOrange"
        }
    }
    catch{
        $ErrorMessage = $_.Exception.Message
        $WPFTextBox_Messages.Text =  ("Couldn't uninstall the PackagingToolDriver / $ErrorMessage")
        $WPFTextBox_Messages.Foreground = "Red"
    }
}


Function Get-PackagingToolDriverStatus{


    try{
    $PackagingToolDriver = Get-WindowsCapability -Online | where  Name -Like "Msix.PackagingTool.Driver*"


    $Name = $PackagingToolDriver.Name
    $State = $PackagingToolDriver.State

    $WPFTextBox_Messages.Text =  ("The state of $Name is $State")
    $WPFTextBox_Messages.Foreground = "Black"


    }
    catch{
        $ErrorMessage = $_.Exception.Message
        $WPFTextBox_Messages.Text =  ("Couldn't get the state / $ErrorMessage")
        $WPFTextBox_Messages.Foreground = "Red"
    }

}


Function Install-MSIX {

    param(
        [Parameter(Mandatory=$true)]
        [String]
        $Path
    )
    
    $CleandUpPath = $Path.replace('"',"")
    
    IF(Test-Path -Path $CleandUpPath){
    
        If((get-item -Path $CleandUpPath).Extension.ToUpper() -eq ".MSIX"){
            $package = get-item -Path $CleandUpPath

            try{
                Add-AppxPackage -Path $package.Fullname -ErrorAction Stop

                $WPFTextBox_Messages.Text =  ("Added package " +$package.Name)
                $WPFTextBox_Messages.Foreground = "Black"

            }catch{
                $ErrorMessage = $_.Exception.Message

                $WPFTextBox_Messages.Text =  ("Failed to add package " +$package.Name +" / $ErrorMessage")
                $WPFTextBox_Messages.Foreground = "Red"
            }
        }
        else{
            $WPFTextBox_Messages.Text =  ("The File seams not to be an MSIX $CleandUpPath")
            $WPFTextBox_Messages.Foreground = "Red"
        }
    }
    else{
        $WPFTextBox_Messages.Text =  ("File not found $CleandUpPath")
        $WPFTextBox_Messages.Foreground = "Red"
    }


}

Function Test-MSIX {

    param(
        [Parameter(Mandatory=$true)]
        [String]
        $Path
    )
    $VBS = new-object -comobject wscript.shell

    $CleandUpPath = $Path.replace('"',"")

    IF(Test-Path -Path $CleandUpPath){


        If((get-item -Path $CleandUpPath).Extension.ToUpper() -eq ".MSIX"){
            $FoundPackages = get-item -Path $CleandUpPath
        }
        else{
            $FoundPackages = (Get-ChildItem -Path $CleandUpPath -Recurse -Filter "*.msix" ) | Sort-Object
        }


        If($FoundPackages.count -ne 1){
            $WPFTextBox_Messages.Text =  ("In the new window, select all the MSIX packages you want to test")
            $WPFTextBox_Messages.Foreground = "Black"

            $SelctedPackages = $FoundPackages | Select-Object -Property Name,FullName |Out-GridView -Title ("Select MSIX files you want to test ("+$FoundPackages.count+")" )-OutputMode Multiple
        }
        else{
            $SelctedPackages = $FoundPackages
        }
        
        ForEach($package in $SelctedPackages){
            $Message = ""

            try{
                $AppxPackagesBefore = get-AppxPackage
                Add-AppxPackage -Path $package.Fullname -ErrorAction Stop
                $AppxPackagesAfter = get-AppxPackage
                $Difference = Compare-Object -ReferenceObject $AppxPackagesBefore -DifferenceObject $AppxPackagesAfter
            
                If($Difference.InputObject.count -gt 0){

                    $ManifestPath = ($Difference.InputObject.InstallLocation +"\AppxManifest.xml")
                    
                    [xml]$ManifestContent = Get-Content -Path $ManifestPath


                    $Applications = $ManifestContent.Package.Applications.Application
            
                    ForEach($Application in $Applications){
                        [String]$Message += ("`n - " + $Application.VisualElements.DisplayName)
                    }

                    $WPFTextBox_Messages.Text =  ("Added MSIX " +$package.Name)
                    $WPFTextBox_Messages.Foreground = "Black"

                    $VBS.popup(("Added MSIX " +$package.Name +" `n`nReady for Test. `n`nThe following things changed in the Startmenu: `n" +$Message),0,"OK",64)
                    $VBS.popup(("Only klick on OK when your done testing " +$package.Name +"`nIt will be uninstalled now!"),0,"OK",64)
                
                    try{
                        $Difference.InputObject | Remove-AppxPackage -ErrorAction Stop

                        $WPFTextBox_Messages.Text =  ("Removed MSIX " +$package.Name)
                        $WPFTextBox_Messages.Foreground = "Black"

                        $VBS.popup(("Removed MSIX " +$package.Name),0,"OK",64)

                    }catch{
                            $ErrorMessage = $_.Exception.Message

                            $WPFTextBox_Messages.Text =  ("Failed to remove package " +$package.Name +" / $ErrorMessage")
                            $WPFTextBox_Messages.Foreground = "Red"

                            $VBS.popup(("Failed to remove MSIX " +$package.Name +"`n`n$ErrorMessage"),0,"Error",16)
                    }

                }
                else{

                    $WPFTextBox_Messages.Text =  ("I think the MSIX " +$package.Name +" was already installed")
                    $WPFTextBox_Messages.Foreground = "Black"

                    $VBS.popup(("I think the MSIX " +$package.Name +" was already installed"),0,"OK",64)
                }

            }catch{
                $ErrorMessage = $_.Exception.Message

                $WPFTextBox_Messages.Text =  ("Failed to add MSIX " +$package.Name +" / $ErrorMessage")
                $WPFTextBox_Messages.Foreground = "Red"

                $VBS.popup(("Failed to add MSIX " +$package.Name +"`n`n$ErrorMessage"),0,"Error",16)
            }

        }

        $WPFTextBox_Messages.Text =  ("All selected " +$SelctedPackages.count +" Packages are done.")
        $WPFTextBox_Messages.Foreground = "Black"
    }
    else{
        $WPFTextBox_Messages.Text =  ("File or Folder not found $CleandUpPath")
        $WPFTextBox_Messages.Foreground = "Red"
    }

}

Function uninstall-App{
    $SelectedMSIX = $null
    $SelectedMSIX = $WPFListView_MSIXPackages.SelectedItem.Name
    $SelectedMSIXInstallLocation= $WPFListView_MSIXPackages.SelectedItem.InstallLocation

    If($MSIXData){
        IF($SelectedMSIX){
            $WPFTextBox_Messages.Text =  "The selected MSIX is $SelectedMSIX"
            $WPFTextBox_Messages.Foreground = "Black"
            try{
                Get-AppxPackage | where InstallLocation -EQ $SelectedMSIXInstallLocation | Remove-AppxPackage
                Get-InstalledMSIX
                $WPFTextBox_Messages.Text =  "Uninstalled the MSIX $SelectedMSIX"
                $WPFTextBox_Messages.Foreground = "Black"
            }
            catch{
                $ErrorMessage = $_.Exception.Message

                $WPFTextBox_Messages.Text ="Failed to uninstall the App  $SelectedMSIX / $ErrorMessage"
                $WPFTextBox_Messages.Foreground = "RED"
            }
        }
        else{
            $WPFTextBox_Messages.Text =  "You need to select a MSIX first"
            $WPFTextBox_Messages.Foreground = "Black"
        }
    }
    else{
        $WPFTextBox_Messages.Text = "You need to get the Software first and then select an MSIX!"
        $WPFTextBox_Messages.Foreground = "Black"
    }

}

Function Open-UserData{
    $SelectedMSIX = $null
    $SelectedMSIX = $WPFListView_MSIXPackages.SelectedItem.Name
    $SelectedMSIXInstallLocation= $WPFListView_MSIXPackages.SelectedItem.InstallLocation

    $SelectedMSIXInstallPublisherHash = $SelectedMSIXInstallLocation.Substring($SelectedMSIXInstallLocation.LastIndexOf("_")+1,$SelectedMSIXInstallLocation.Length -$SelectedMSIXInstallLocation.LastIndexOf("_")-1)
    $UserDataLocation = "$env:LOCALAPPDATA\Packages\$SelectedMSIX"+"_"+$SelectedMSIXInstallPublisherHash


    If($MSIXData){
        IF($SelectedMSIX){
            $WPFTextBox_Messages.Text =  "The selected MSIX is $UserDataLocation"
            $WPFTextBox_Messages.Foreground = "Black"
            explorer.exe $UserDataLocation


        }
        else{
            $WPFTextBox_Messages.Text =  "You need to select a MSIX first"
            $WPFTextBox_Messages.Foreground = "Black"
        }
    }
    else{
        $WPFTextBox_Messages.Text = "You need to get the Software first and then select an MSIX!"
        $WPFTextBox_Messages.Foreground = "Black"
    }

}

Function Open-InstallLocation{
    $SelectedMSIX = $null
    $SelectedMSIX = $WPFListView_MSIXPackages.SelectedItem.Name
    $SelectedMSIXInstallLocation= $WPFListView_MSIXPackages.SelectedItem.InstallLocation

    If($MSIXData){
        IF($SelectedMSIX){
            $WPFTextBox_Messages.Text =  "The selected MSIX is $SelectedMSIX"
            $WPFTextBox_Messages.Foreground = "Black"
            explorer.exe $SelectedMSIXInstallLocation


        }
        else{
            $WPFTextBox_Messages.Text =  "You need to select a MSIX first"
            $WPFTextBox_Messages.Foreground = "Black"
        }
    }
    else{
        $WPFTextBox_Messages.Text = "You need to get the Software first and then select an MSIX!"
        $WPFTextBox_Messages.Foreground = "Black"
    }

}

Function Get-InstalledMSIX{

        $Script:MSIXData = ""
        $Script:MSIXData = @()


    If($WPFCheckBox_EnterpriseSigned.IsChecked){

        $InstalledApps = Get-AppxPackage | Where SignatureKind -EQ "Enterprise" |Select-Object -Property Name,Version,Publisher,InstallLocation  

    }
    else{

        $InstalledApps = Get-AppxPackage | Select-Object -Property Name,Version,Publisher,InstallLocation  

    }

    ForEach($InstalledApp in $InstalledApps){

        $Name = $InstalledApp.Name
        $Publisher = $InstalledApp.Publisher
        $Version = $InstalledApp.Version
        $InstallLocation = $InstalledApp.InstallLocation

        #$InstallDate = [Datetime]::ParseExact($InstalledSoftware.InstallDate,(Get-culture).DateTimeFormat.ShortDatePattern +" " +(Get-culture).DateTimeFormat.LongTimePattern,$null)

        $SoftwareDetail = New-Object PSObject
        $SoftwareDetail | Add-Member -Name "Name" -MemberType NoteProperty -Value $Name
        $SoftwareDetail | Add-Member -Name "Version" -MemberType NoteProperty -Value $Version
        $SoftwareDetail | Add-Member -Name "Publisher" -MemberType NoteProperty -Value $Publisher
        $SoftwareDetail | Add-Member -Name "InstallLocation" -MemberType NoteProperty -Value $InstallLocation

        $Script:MSIXData += $SoftwareDetail
    }


    $WPFListView_MSIXPackages.ItemsSource = $MSIXData
    $WPFTextBox_Messages.Text =  "Got the list of Appx/MSIX installed software"
    $WPFTextBox_Messages.Foreground = "Black"

}

Function Use-Filter{

    If($MSIXData){        
        If($WPFTexBox_Filter.Text){
            $WPFTextBox_Messages.Text =  "Setting filter to " +$WPFTexBox_Filter.Text
            $WPFTextBox_Messages.Foreground = "Black"

        }
        else{
            $WPFTextBox_Messages.Text =  "Removing Filter"
            $WPFTextBox_Messages.Foreground = "Black"
        }

        #Use Filter
        $filter = $WPFTexBox_Filter.Text
        $viewMSIXPackages = [System.Windows.Data.CollectionViewSource]::GetDefaultView($WPFListView_MSIXPackages.ItemsSource)
        $viewMSIXPackages.Filter = {param ($item) $item -match $filter}
           
        $viewMSIXPackages.Refresh()

    }
    else{
        $WPFTextBox_Messages.Text = "You need to get the Software first!"
        $WPFTextBox_Messages.Foreground = "Black"
    }
}

Function Import-CertToTrustedPeople{

    #Check if user is Admin
    If ([bool](([System.Security.Principal.WindowsIdentity]::GetCurrent()).groups -match "S-1-5-32-544")) {
        #If he is an Admin
        $CertToInstall = $WPFTextBox_SelectedCert.Text
        If(-!$CertToInstall){
            $WPFTextBox_Messages.Text ="No Cert provided"
            $WPFTextBox_Messages.Foreground = "RED"
            Return
        }
        else{
            If(Test-Path -Path $CertToInstall){

                try{
                 Import-Certificate -FilePath $CertToInstall -CertStoreLocation cert:\LocalMachine\TrustedPeople
                    $WPFTextBox_Messages.Text ="Succesfully installed Cert to LocalMachine\TrustedPeople $CertToInstall"
                    $WPFTextBox_Messages.Foreground = "Black"

                }
                catch{
                    $ErrorMessage = $_.Exception.Message

                    $WPFTextBox_Messages.Text ="Failed to install the cert File / $ErrorMessage"
                    $WPFTextBox_Messages.Foreground = "RED"
                }

            }
            else{
                $WPFTextBox_Messages.Text ="The Path $CertToInstall is not valid"
                $WPFTextBox_Messages.Foreground = "RED"
                Return
            }

        }
     }
    else{
         #If he is NO Admin
        $WPFTextBox_Messages.Text = "You are not an administrator. This Action requieres the tool to be run as an administrator!"
        $WPFTextBox_Messages.Foreground = "Red"
        Return
    }

}

Function Select-Cert{
    $WPFTextBox_SelectedCert.Text = $null

    #Select Cer File
    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null    
    $OpenFileDialogCER = New-Object System.Windows.Forms.OpenFileDialog
    #$OpenFileDialogCER.initialDirectory = ((get-item $ThisScriptParentPath).Parent).FullName
    $OpenFileDialogCER.filter = "CER (*.cer)| *.cer"
    $OpenFileDialogCER.Title = "Select the Cer File that you want to install"
    $OpenFileDialogCER.ShowDialog() | Out-Null
    $CER=  $OpenFileDialogCER.FileName
    $CERFileName = $OpenFileDialogCER.SafeFileName

    If($CER){
        If($CERFileName.ToUpper().Contains(".CER") -gt 0){
            $WPFTextBox_SelectedCert.Text = $CER
            $WPFTextBox_Messages.Text = "$CERFileName selected. This is a valid CER-File"
            $WPFTextBox_Messages.Foreground = "Black"
        }
        else{
            $WPFTextBox_Messages.Text = "$CERFileName is not a valid CER-File! Select something else."
            $WPFTextBox_Messages.Foreground = "Red"
        }
    }
    else{
        $WPFTextBox_Messages.Text = "Nothing selected!"
        $WPFTextBox_Messages.Foreground = "Red"
    }

}

Function Creat-SelfSignedCert{
    $YourPublisher = $WPFTextBox_NewCertPublisherName.Text #Name for the Publisher
    $FriendlyName = $WPFTextBox_NewCertFriendlyName.Text #What ever friendly name you like
    $password = $WPFPasswordBox_NewCertPassword.Password #Remember that one, you need it again when creating your MSIX
    $FolderToExportCert = $WPFTextBox_NewCertFolderName.Text #The folder already needs to exist.
    $CertName= $WPFTextBox_NewCertFileName.Text #The Filename your Cert should have


    If(-!$YourPublisher){
        $WPFTextBox_Messages.Text ="No Publishername provided"
        $WPFTextBox_Messages.Foreground = "RED"
        Return
    }
    else{
        If(-!$YourPublisher.indexof("CN=")-gt 0){
           $YourPublisher = "CN=$YourPublisher"
        }
    }


    If(-!$FriendlyName){
        $WPFTextBox_Messages.Text ="No Friendlyname provided"
        $WPFTextBox_Messages.Foreground = "RED"
        Return
    }

    If(-!$password){
        $WPFTextBox_Messages.Text ="No password provided"
        $WPFTextBox_Messages.Foreground = "RED"
        Return
    }

    If(-!$FolderToExportCert){
        $WPFTextBox_Messages.Text ="No Export Folder provided"
        $WPFTextBox_Messages.Foreground = "RED"
        Return
    }

    If(-!$CertName){
        $WPFTextBox_Messages.Text ="No Filename provided"
        $WPFTextBox_Messages.Foreground = "RED"
        Return
    }


    try{

        $PFXName= ($FolderToExportCert +'\' +$CertName +'.pfx')
        $CertName= ($FolderToExportCert +'\' +$CertName +'.cer')


        If(Test-Path $PFXName){
            $WPFTextBox_Messages.Text ="The File $PFXName already exists. I will do nothing. Choose an other name or remove the file first."
            $WPFTextBox_Messages.Foreground = "RED"
            Return
        }
        else{
            If(Test-Path $CertName){
                $WPFTextBox_Messages.Text ="The File $CertName already exists. I will do nothing. Choose an other name or remove the file first."
                $WPFTextBox_Messages.Foreground = "RED"
                Return
            }
            else{
                $cert = New-SelfSignedCertificate -Type Custom -Subject $YourPublisher -KeyUsage DigitalSignature -FriendlyName $FriendlyName -CertStoreLocation 'Cert:\CurrentUser\my'
                $pwd = ConvertTo-SecureString -String $password -Force -AsPlainText 
                $cert | Export-PfxCertificate -FilePath $PFXName -Password $pwd
                $cert | Export-Certificate -Type CERT -FilePath $CertName
                remove-item $cert.PSPath

                $WPFTextBox_Messages.Text ="Created Certfile $PFXName"
                $WPFTextBox_Messages.Foreground = "Black"
            }        
        }

    }
    catch{

        $ErrorMessage = $_.Exception.Message

        $WPFTextBox_Messages.Text ="Failed to creat File $PFXName / $ErrorMessage"
        $WPFTextBox_Messages.Foreground = "RED"
    }
}

Function Select-Folder{
    Add-Type -AssemblyName System.Windows.Forms
    $FolderBrowser = New-Object System.Windows.Forms.FolderBrowserDialog
    [void]$FolderBrowser.ShowDialog()
    $FolderBrowser.SelectedPath
}

Function Select-MSIX {

    #Select MSIX File
    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null    
    $OpenFileDialogMSIX = New-Object System.Windows.Forms.OpenFileDialog
    #$OpenFileDialogMSIX.initialDirectory = ((get-item $ThisScriptParentPath).Parent).FullName
    $OpenFileDialogMSIX.filter = "MSIX (*.msix)| *.msix"
    $OpenFileDialogMSIX.Title = "Select the MSIX File that you want to use"
    $OpenFileDialogMSIX.ShowDialog() | Out-Null
    $MSIX=  $OpenFileDialogMSIX.FileName
    $MSIXFileName = $OpenFileDialogMSIX.SafeFileName

    If($MSIX){
        If($MSIXFileName.ToUpper().Contains(".MSIX") -gt 0){
            $WPFTextBox_Messages.Text = "$MSIXFileName selected. This is a valid MSIX-File"
            $WPFTextBox_Messages.Foreground = "Black"
            $MSIX
        }
        else{
            $WPFTextBox_Messages.Text = "$MSIXFileName is not a valid MSIX-File! Select something else."
            $WPFTextBox_Messages.Foreground = "Red"
        }
    }
    else{
        $WPFTextBox_Messages.Text = "Nothing selected!"
        $WPFTextBox_Messages.Foreground = "Red"
    }


}

#endregion

#region Actions

#Installed Actions
$WPFButton_GetSoftware.Add_Click({

    If($WPFListView_MSIXPackages.ItemsSource){
        $WPFListView_MSIXPackages.ItemsSource.Clear()
        $WPFListView_MSIXPackages.Items.Refresh()
    }

    Get-InstalledMSIX

})


$WPFButton_Filter.Add_Click({
    Use-Filter
})

$WPFButton_OpenInstallLocation.Add_Click({
    Open-InstallLocation
})

$WPFButton_OpenUserData.Add_Click({
    Open-UserData
})

$WPFButton_Uninstall.Add_Click({
    uninstall-App
})


#Cert Acrtions

#Open Browse CertExportFolder Button
$WPFButton_SelectCertStoreFolder.Add_Click({
    $SelectedFolder = $null
    $SelectedFolder = Select-Folder

    If($SelectedFolder){
        $WPFTextBox_Messages.Text ="The Folder $SelectedFolder was selected"
        $WPFTextBox_NewCertFolderName.Text = $SelectedFolder
        $WPFTextBox_Messages.Foreground = "Black"

    }
    else{
        $WPFTextBox_Messages.Text ="No Folder was selected"
        $WPFTextBox_Messages.Foreground = "DarkOrange"

    }
})

$WPFButton_CreatCert.Add_Click({
    Creat-SelfSignedCert
})

$WPFButton_SelectCertToImport.Add_Click({
    Select-Cert
})


$WPFButton_InstallCert.Add_Click({
    Import-CertToTrustedPeople
})

#Test Actions

$WPFButton_TestSelectInstallMSIXFile.Add_Click({

    
    $WPFTextBox_TestSelectedInstallMSIX.Text = $null

    $SelectedMSIX = Select-MSIX

    $WPFTextBox_TestSelectedInstallMSIX.Text = $SelectedMSIX

})


$WPFButton_TestOnlyInstall.Add_Click({

    $SelectedPath = $WPFTextBox_TestSelectedInstallMSIX.Text
    If($SelectedPath){
        Install-MSIX -Path $SelectedPath
    }
    else{
        $WPFTextBox_Messages.Text = "Nothing selected!"
        $WPFTextBox_Messages.Foreground = "Red"
    }
})


$WPFButton_TestSelectSingleMSIXFile.Add_Click({
    
    $WPFTextBox_TestSelectedSingleMSIX.Text = $null

    $SelectedMSIX = Select-MSIX

    $WPFTextBox_TestSelectedSingleMSIX.Text = $SelectedMSIX

})

$WPFButton_TestSingleMSIX.Add_Click({


    $SelectedPath = $WPFTextBox_TestSelectedSingleMSIX.Text
    If($SelectedPath){
        Test-MSIX -Path $SelectedPath
    }
    else{
        $WPFTextBox_Messages.Text = "Nothing selected!"
        $WPFTextBox_Messages.Foreground = "Red"
    }
})

$WPFButton_TestSelectMultipleFolder.Add_Click({
    $SelectedFolder = $null
    $SelectedFolder = Select-Folder

    If($SelectedFolder){
        $WPFTextBox_Messages.Text ="The Folder $SelectedFolder was selected"
        $WPFTextBox_TestSelectedMultipleMSIXFolder.Text = $SelectedFolder
        $WPFTextBox_Messages.Foreground = "Black"

    }
    else{
        $WPFTextBox_Messages.Text ="No Folder was selected"
        $WPFTextBox_Messages.Foreground = "DarkOrange"

    }
})

$WPFButton_TestMultipleMSIX.Add_Click({

    $SelectedPath = $WPFTextBox_TestSelectedMultipleMSIXFolder.Text
    If($SelectedPath){
        Test-MSIX -Path $SelectedPath
    }
    else{
        $WPFTextBox_Messages.Text = "Nothing selected!"
        $WPFTextBox_Messages.Foreground = "Red"
    }
})



# Packaging Tool Driver
$WPFButton_PackagingToolDriverCheckStatus.Add_Click({
    Get-PackagingToolDriverStatus
})

$WPFButton_PackagingToolDriverInstall.Add_Click({
    $WPFTextBox_Messages.Text =  ("Installing will take some time. Please be patient.")
    $WPFTextBox_Messages.Foreground = "Black"
    Install-PackagingToolDriver
})

$WPFButton_PackagingToolDriverUninstall.Add_Click({
    Uninstall-PackagingToolDriver
})



#endregion


#===========================================================================
# Shows the form
#===========================================================================
$Form.ShowDialog() | out-null


