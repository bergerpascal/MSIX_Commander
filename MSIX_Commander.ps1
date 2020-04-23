$ScriptVersion = "1.0.7.3"
# Add required assemblies for icon
Add-Type -AssemblyName PresentationFramework, System.Drawing, System.Windows.Forms, WindowsFormsIntegration

#region GUI
#ERASE ALL THIS AND PUT XAML BELOW between the @" "@
$inputXML = @"
<Window x:Class="MSIX_Commander.MainWindow"
        xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
        xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
        xmlns:d="http://schemas.microsoft.com/expression/blend/2008"
        xmlns:mc="http://schemas.openxmlformats.org/markup-compatibility/2006"
        xmlns:local="clr-namespace:MSIX_Commander"
        mc:Ignorable="d"
        Title="MSIX Commander" Width="800" Height="600">
    <Window.TaskbarItemInfo>
        <TaskbarItemInfo/>
    </Window.TaskbarItemInfo>
    <Grid>
        <Grid.ColumnDefinitions>
            <ColumnDefinition/>
        </Grid.ColumnDefinitions>
        <Label x:Name="Label_Messages" Content="Messages:" HorizontalAlignment="Left" Height="27" VerticalAlignment="Top" Width="64" Margin="0,-4,0,0"/>
        <TextBox x:Name="TextBox_Messages" Height="37" Margin="69,4,10,0" TextWrapping="Wrap" VerticalAlignment="Top" IsEnabled="False"/>
        <TabControl Margin="0,46,10,23.5">
            <TabItem Header="Installed">
                <Grid Background="#FFE5E5E5">
                    <Grid.RowDefinitions>
                        <RowDefinition Height="107*"/>
                        <RowDefinition Height="339*"/>
                        <RowDefinition Height="30*"/>
                    </Grid.RowDefinitions>
                    <ListView x:Name="ListView_MSIXPackages" Margin="0,23.5,0,10" HorizontalAlignment="Left" MinWidth="635" MinHeight="330" Grid.RowSpan="2" Grid.Row="1">
                        <ListView.View>
                            <GridView>
                                <GridViewColumn Header="Name" DisplayMemberBinding ="{Binding Name}"/>
                                <GridViewColumn Header="Version" DisplayMemberBinding ="{Binding Version}"/>
                                <GridViewColumn Header="PackageType" DisplayMemberBinding ="{Binding PackageType}"/>
                                <GridViewColumn Header="Publisher" DisplayMemberBinding ="{Binding Publisher}"/>
                                <GridViewColumn Header="InstallLocation" DisplayMemberBinding ="{Binding InstallLocation}"/>
                                <GridViewColumn Header="Modifications" DisplayMemberBinding ="{Binding Dependencies}"/>
                            </GridView>
                        </ListView.View>
                    </ListView>
                    <Button x:Name="Button_GetSoftware" Content="Get Software" HorizontalAlignment="Left" Margin="0,32,0,0" VerticalAlignment="Top" Width="93"/>
                    <Button x:Name="Button_Filter" Content="Filter" HorizontalAlignment="Left" Margin="560,31,0,0" VerticalAlignment="Top" Width="75"/>
                    <TextBox x:Name="TexBox_Filter" HorizontalAlignment="Left" Height="23" Margin="334,29,0,0" TextWrapping="Wrap" Text="" VerticalAlignment="Top" Width="221"/>
                    <Label x:Name="Label_Filter" Content="Filter:" HorizontalAlignment="Left" Height="23" VerticalAlignment="Top" Width="46" Margin="298,29,0,0"/>
                    <CheckBox x:Name="CheckBox_EnterpriseSigned" Content="Only Enterprise/Developer signed" HorizontalAlignment="Left" Margin="99,34,0,0" VerticalAlignment="Top" IsChecked="True"/>
                    <Button x:Name="Button_OpenInstallLocation" Content="Open Installlocation" HorizontalAlignment="Left" VerticalAlignment="Top" Width="114" Margin="0,58,0,0"/>
                    <Button x:Name="Button_OpenUserData" Content="Open Userdata" HorizontalAlignment="Left" Margin="119,58,0,0" Width="114" Height="20" VerticalAlignment="Top"/>
                    <Label x:Name="Label_Installed" Content="Work with installed MSIX Packages" HorizontalAlignment="Left" FontWeight="Bold" Height="26" VerticalAlignment="Top" Margin="-2,0,0,0"/>
                    <Button x:Name="Button_Uninstall" Content="Uninstall" HorizontalAlignment="Left" Margin="357,58,0,0" VerticalAlignment="Top" Width="114"/>
                    <Button x:Name="Button_Open_Manifest" Content="Open Manifest" HorizontalAlignment="Left" Margin="238,58,0,0" VerticalAlignment="Top" Width="114"/>
                    <Button x:Name="Button_OpenCMD" Content="Open CMD" HorizontalAlignment="Left" VerticalAlignment="Top" Width="114" Margin="0,105,0,0" Grid.RowSpan="2"/>
                    <Button x:Name="Button_OpenRegedit" Content="Open Regedit*" HorizontalAlignment="Left" VerticalAlignment="Top" Width="114" Margin="119,105,0,0" Grid.RowSpan="2"/>
                    <Button x:Name="Button_OpenNotepad" Content="Open Notepad" HorizontalAlignment="Left" VerticalAlignment="Top" Width="114" Margin="238,105,0,0" Grid.RowSpan="2"/>
                    <TextBlock x:Name="TextBlock_OpenTools" HorizontalAlignment="Left" TextWrapping="Wrap" Text="Here you can open the following tools inside the MSIX container. This requires developer mode to be enabled on your machine." VerticalAlignment="Top" Margin="0,86,0,0"/>
                    <TextBlock x:Name="TextBlock_OpenTools_Copy" HorizontalAlignment="Left" TextWrapping="Wrap" VerticalAlignment="Top" Margin="357,106,0,0" Height="21" Width="382" Text="Notepad is a workaround since explorer is not working right now." Grid.RowSpan="2"/>
                    <Button x:Name="Button_Start" Content="Start" HorizontalAlignment="Left" Margin="476,58,0,0" VerticalAlignment="Top" Width="114"/>
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
                    <Label x:Name="Label_ImportCert" Content="Import a selfsigned cert from a *.cert File*" HorizontalAlignment="Left" VerticalAlignment="Top" FontWeight="Bold" Margin="-2,203,0,0"/>
                    <Button x:Name="Button_InstallCert" Content="Install cert*" Margin="326,285,0,0" VerticalAlignment="Top" Height="33" Width="128" HorizontalAlignment="Left"/>
                    <Button x:Name="Button_SelectCertToImport" Content="Browse" Margin="0,250,7,0" VerticalAlignment="Top" Height="26" HorizontalAlignment="Right" Width="48"/>
                    <Label x:Name="Label_SelectedCert" Content="Selected Cert:" HorizontalAlignment="Left" Margin="3,250,0,0" VerticalAlignment="Top"/>
                    <TextBox x:Name="TextBox_SelectedCert" Height="26" Margin="100,250,61,0" TextWrapping="Wrap" VerticalAlignment="Top"/>
                    <TextBlock x:Name="TexBlock_InstallCert" HorizontalAlignment="Left" Margin="3,229,0,0" TextWrapping="Wrap" Text="This will install the cert to the TrustedPeople Store of the local machine. This requires Administrator rights." VerticalAlignment="Top" Height="27" Width="758"/>
                    <Label x:Name="Label_ImportCertFromMSIX" Content="Import a selfsigned cert from a *.MSIX File*" HorizontalAlignment="Left" VerticalAlignment="Top" FontWeight="Bold" Margin="-2,319,0,0"/>
                    <Button x:Name="Button_InstallCertFromMSIX" Content="Install cert*" Margin="326,401,0,0" VerticalAlignment="Top" Height="33" Width="128" HorizontalAlignment="Left"/>
                    <Button x:Name="Button_SelectCertToImportFromMSIX" Content="Browse" Margin="0,366,7,0" VerticalAlignment="Top" Height="26" HorizontalAlignment="Right" Width="48"/>
                    <Label x:Name="Label_SelectedCertFromMSIX" Content="Selected MSIX:" HorizontalAlignment="Left" Margin="3,366,0,0" VerticalAlignment="Top"/>
                    <TextBox x:Name="TextBox_SelectedCertFromMSIX" Height="26" Margin="100,366,61,0" TextWrapping="Wrap" VerticalAlignment="Top"/>
                    <TextBlock x:Name="TexBlock_InstallCertFromMSIX" HorizontalAlignment="Left" Margin="3,345,0,0" TextWrapping="Wrap" Text="This will install the cert ditrectly from an MSIX to the TrustedPeople Store of the local machine. This requires Administrator rights." VerticalAlignment="Top" Height="27" Width="758"/>
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
            <TabItem x:Name="Tab_PackagingMachien" Header="PackagingMachine">
                <Grid Background="#FFE5E5E5" Margin="0,0,0,-183">
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition/>
                    </Grid.ColumnDefinitions>
                    <Button x:Name="Button_PackagingToolDriverCheckStatus" Content="Check state*" Margin="141,74,0,0" VerticalAlignment="Top" Height="33" Width="128" HorizontalAlignment="Left"/>
                    <Label x:Name="Label_PackagingToolDriverCheck" Content="Check Packaging Tool Driver state" HorizontalAlignment="Left" VerticalAlignment="Top" FontWeight="Bold" Margin="0,-5,0,0"/>
                    <TextBlock x:Name="TexBlock_PackagingToolDriverCheck" HorizontalAlignment="Left" Margin="6,21,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Height="36" Width="758" Text="This will let you check the state of the Packaging Tool Driver FeatureOnDemand (FOD). You can also install and uninstall it. Installing will take some time, please be patient. This requires Administrator rights. It only works with Windows 10 1809 or higher."/>
                    <Button x:Name="Button_PackagingToolDriverInstall" Content="Install from WindowsUpdate*" Margin="307,74,0,0" VerticalAlignment="Top" Height="33" HorizontalAlignment="Left" Width="164"/>
                    <Button x:Name="Button_PackagingToolDriverUninstall" Content="Uninstall*" Margin="506,74,0,0" VerticalAlignment="Top" Height="33" Width="128" HorizontalAlignment="Left"/>
                    <Label x:Name="Label_StopServices" Content="Stop services" HorizontalAlignment="Left" VerticalAlignment="Top" FontWeight="Bold" Margin="0,107,0,0"/>
                    <TextBlock x:Name="TexBlock_StopServices" HorizontalAlignment="Left" Margin="6,133,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Height="48" Width="758"><Run Text="The following services will be stopped and disable"/><Run Text="d"/><Run Text=" if they are present on your system:"/><LineBreak/><Run Text="Diagnostic Policy Service, Offline Files, Windows Search, Windows Update, SMS-Agent-Host, PLRestartMgrService"/><LineBreak/><Run Text="This requires Administrator rights. "/></TextBlock>
                    <Button x:Name="Button_StopServices" Content="Stop services*" Margin="307,186,0,0" VerticalAlignment="Top" Height="33" HorizontalAlignment="Left" Width="164"/>
                    <Label x:Name="Label_Sideloading" Content="Change Sideloading Status" HorizontalAlignment="Left" VerticalAlignment="Top" FontWeight="Bold" Margin="0,225,0,0"/>
                    <TextBlock x:Name="TexBlock_Sideloading" HorizontalAlignment="Left" Margin="6,251,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Height="22" Width="758"><Run Text="Here you can enable or disable sideloading."/><Run Text=" "/><Run Text="This requires Administrator rights."/><LineBreak/><Run Text=""/><LineBreak/><Run Text=""/></TextBlock>
                    <Button x:Name="Button_Change_SidelaodingStatus" Content="" Margin="161,300,0,0" VerticalAlignment="Top" Height="33" HorizontalAlignment="Left" Width="164"/>
                    <TextBlock x:Name="TexBlock_CurrentSideloading" HorizontalAlignment="Left" Margin="6,273,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Height="22" Width="179"><Span Foreground="Black" FontSize="12" FontFamily="Segoe UI"><Run Text="The current sideloading status is:"/></Span><Span Foreground="Black" FontSize="12" FontFamily="Segoe UI"><LineBreak/></Span><LineBreak/><Run Text=""/><LineBreak/><Run Text=""/></TextBlock>
                    <TextBlock x:Name="TexBlock_CurrentSideloading_Status" HorizontalAlignment="Left" Margin="190,273,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Height="22" Width="179" FontWeight="Bold" Foreground="#FFFF0600"/>
                    <Button x:Name="Button_EnableDevMode" Content="Enable Developer Mode*" Margin="438,300,0,0" VerticalAlignment="Top" Height="33" HorizontalAlignment="Left" Width="164"/>
                </Grid>
            </TabItem>
            <TabItem x:Name="Tab_EditManifest" Header="EditManifest">
                <Grid Background="#FFE5E5E5" Margin="0,0,0,-183">
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition/>
                    </Grid.ColumnDefinitions>
                    <Button x:Name="Button_EditManifest_SingelFile_Edit" Content="Edit Manifest" Margin="325,402,0,0" VerticalAlignment="Top" Height="26" HorizontalAlignment="Left" Width="128"/>
                    <Label x:Name="Label_EditManifest_FileOrFolder" Content="What to edit" HorizontalAlignment="Left" VerticalAlignment="Top" FontWeight="Bold" Margin="-2,281,0,0"/>
                    <TextBlock x:Name="TexBlock_EditManifest_FileOrFolder" HorizontalAlignment="Left" Margin="3,307,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Height="37" Width="757"><Run Text="This will let you select one single MSIX file or a folder containing MSIX Files"/><Run Text="."/><Run Text=" "/><Run Text="If you select a folder, the tool will also look in it's subfolders for MSIX Files."/></TextBlock>
                    <Button x:Name="Button_EditManifest_SingelFile_Browse" Content="Select single MSIX" Margin="233,370,0,0" VerticalAlignment="Top" Height="26" HorizontalAlignment="Left" Width="113"/>
                    <Label x:Name="Label_EditManifest_SingelFile_MSIX" Content="File or Folder:" HorizontalAlignment="Left" Margin="-2,341,0,0" VerticalAlignment="Top"/>
                    <TextBox x:Name="TextBox_EditManifest_SingelFile_SelectedMSIX" Height="20" Margin="104,345,71,0" TextWrapping="Wrap" VerticalAlignment="Top" MaxHeight="20"/>
                    <Button x:Name="Button_EditManifest_SelectCertToImport" Content="Browse" Margin="0,162,17,0" VerticalAlignment="Top" Height="26" HorizontalAlignment="Right" Width="48"/>
                    <TextBox x:Name="TextBox_EditManifest_SelectedCert" Height="26" Margin="104,162,71,0" TextWrapping="Wrap" VerticalAlignment="Top"/>
                    <TextBlock x:Name="TexBlock_EditManifest_Cert" HorizontalAlignment="Left" Margin="3,123,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Height="33" Width="758"><Run Text="If you don't specify a PFX certificate and its password, the MSIX will not be signed after the manifest edit and you will not be able to install it. If you don't use the same certificate your original MSIX was signed with, the new MSIX will "/><Run Text="not "/><Run Text="be accepted as an update package for the old one."/></TextBlock>
                    <Label x:Name="Label_EditManifest_Password" Content="Password:" HorizontalAlignment="Left" Margin="-4,192,0,0" VerticalAlignment="Top"/>
                    <PasswordBox x:Name="PasswordBox_EditManifest_CertPassword" HorizontalAlignment="Left" Margin="104,192,0,0" VerticalAlignment="Top" Width="242" Height="26"/>
                    <Label x:Name="Label_EditManifest_DataToChange" Content="Data that should be changed" HorizontalAlignment="Left" VerticalAlignment="Top" FontWeight="Bold" Margin="-2,-1,0,0"/>
                    <Label x:Name="Label_EditManifest_MinVersionTested" Content="MinVersion:" HorizontalAlignment="Left" Margin="-2,74,0,0" VerticalAlignment="Top"/>
                    <TextBox x:Name="TextBox_MinVersionTested" Height="26" Margin="104,74,0,0" TextWrapping="Wrap" VerticalAlignment="Top" HorizontalAlignment="Left" Width="242"/>
                    <Label x:Name="Label_EditManifest_MaxVersionTested" Content="MaxVersionTested:" HorizontalAlignment="Left" Margin="352,74,0,0" VerticalAlignment="Top"/>
                    <TextBox x:Name="TextBox_MaxVersionTested" Height="26" Margin="466,74,71,0" TextWrapping="Wrap" VerticalAlignment="Top"/>
                    <TextBlock x:Name="TexBlock_EditManifest_DataToChange" HorizontalAlignment="Left" Margin="3,21,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Height="56" Width="776"><Run Text="You can specify a new MinVersionTested and MaxVersionTested. If you don't specify one of them, the current value in the Manifest will not be changed. But you must define at least one of them or nothing will happen. "/><LineBreak/><Run Text="Those values must be specified as Windows Build numbers like 10.0.16299.0 or 10.0.17134.0"/></TextBlock>
                    <Label x:Name="Label_EditManifest_Timeserver" Content="Timeserver:" HorizontalAlignment="Left" Margin="394,192,0,0" VerticalAlignment="Top"/>
                    <TextBox x:Name="TextBox_Timeserver" Height="26" Margin="466,192,71,0" TextWrapping="Wrap" VerticalAlignment="Top" Text="http://timestamp.digicert.com"/>
                    <Label x:Name="Label_EditManifest_Cert" Content="Certificate to use" HorizontalAlignment="Left" VerticalAlignment="Top" FontWeight="Bold" Margin="-2,100,0,0"/>
                    <Label x:Name="Label_TexBlock_EditManifest_SelectedCert" Content="Selected PFX Cert:" HorizontalAlignment="Left" Margin="-3,162,0,0" VerticalAlignment="Top"/>
                    <RadioButton x:Name="RadioButton_EditManifest_IncreaseVersion_Yes" Content="Yes" HorizontalAlignment="Left" Margin="104,264,0,0" VerticalAlignment="Top" IsChecked="True" GroupName="IncreaseVersion"/>
                    <RadioButton x:Name="RadioButton_EditManifest_IncreaseVersion_No" Content="No" HorizontalAlignment="Left" Margin="150,264,0,0" VerticalAlignment="Top" GroupName="IncreaseVersion"/>
                    <TextBlock x:Name="TexBlock_EditManifest_IncreaseVersion" HorizontalAlignment="Left" Margin="3,245,0,0" TextWrapping="Wrap" Text="Should the package version number be increased by one, after editing the manifest?" VerticalAlignment="Top" Height="17" Width="758"/>
                    <Label x:Name="Label_EditManifest_IncreaseVersion1" Content="Increase Version:" HorizontalAlignment="Left" Margin="-2,256,0,0" VerticalAlignment="Top"/>
                    <Label x:Name="Label_EditManifest_IncreaseVersion" Content="Increase the Package Version" HorizontalAlignment="Left" VerticalAlignment="Top" FontWeight="Bold" Margin="-2,222,0,0"/>
                    <Button x:Name="Button_EditManifest_SelectFolder" Content="Select Folder" Margin="424,370,0,0" VerticalAlignment="Top" Height="26" HorizontalAlignment="Left" Width="124"/>
                </Grid>
            </TabItem>
            <TabItem x:Name="Tab_ChangeSignature" Header="ChangeSignature">
                <Grid Background="#FFE5E5E5" Margin="0,0,0,-183">
                    <Grid.ColumnDefinitions>
                        <ColumnDefinition/>
                    </Grid.ColumnDefinitions>
                    <Button x:Name="Button_ChangeSignature" Content="Change Signature" Margin="325,244,0,0" VerticalAlignment="Top" Height="26" HorizontalAlignment="Left" Width="128"/>
                    <Label x:Name="Label_ChangeSignature_FileOrFolder" Content="What to edit" HorizontalAlignment="Left" VerticalAlignment="Top" FontWeight="Bold" Margin="-2,123,0,0"/>
                    <TextBlock x:Name="TexBlock_ChangeSignature_FileOrFolder" HorizontalAlignment="Left" Margin="3,149,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Height="37" Width="757"><Run Text="This will let you select one single MSIX file or a folder containing MSIX Files"/><Run Text="."/><Run Text=" "/><Run Text="If you select a folder, the tool will also look in it's subfolders for MSIX Files."/></TextBlock>
                    <Button x:Name="Button_ChangeSignature_SingelFile_Browse" Content="Select single MSIX" Margin="233,212,0,0" VerticalAlignment="Top" Height="26" HorizontalAlignment="Left" Width="113"/>
                    <Label x:Name="Label_ChangeSignature_SingelFile_MSIX" Content="File or Folder:" HorizontalAlignment="Left" Margin="-2,183,0,0" VerticalAlignment="Top"/>
                    <TextBox x:Name="TextBox_ChangeSignature_SelectedMSIX" Height="20" Margin="104,187,71,0" TextWrapping="Wrap" VerticalAlignment="Top" MaxHeight="20"/>
                    <Button x:Name="Button_ChangeSignature_SelectCertToImport" Content="Browse" Margin="0,64,17,0" VerticalAlignment="Top" Height="26" HorizontalAlignment="Right" Width="48"/>
                    <TextBox x:Name="TextBox_ChangeSignature_SelectedCert" Height="26" Margin="104,64,71,0" TextWrapping="Wrap" VerticalAlignment="Top"/>
                    <TextBlock x:Name="TexBlock_ChangeSignature_Cert" HorizontalAlignment="Left" Margin="3,25,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Height="33" Width="758"><Run Text="Your MSIX will be signed with this certificate. This will also change the publisher inside your manifest file to the subject of this certificate."/><LineBreak/><Run Text="You must be aware, that changing the Publisher of the MSIX will break Updates and Modification Packages."/></TextBlock>
                    <Label x:Name="Label_ChangeSignature_Password" Content="Password:" HorizontalAlignment="Left" Margin="-4,94,0,0" VerticalAlignment="Top"/>
                    <PasswordBox x:Name="PasswordBox_ChangeSignatureCertPassword" HorizontalAlignment="Left" Margin="104,94,0,0" VerticalAlignment="Top" Width="242" Height="26"/>
                    <TextBox x:Name="TextBox_ChangeSignature_Timeserver" Height="26" Margin="466,94,71,0" TextWrapping="Wrap" VerticalAlignment="Top" Text="http://timestamp.digicert.com"/>
                    <Label x:Name="Label_ChangeSignature_Cert" Content="Certificate to use" HorizontalAlignment="Left" VerticalAlignment="Top" FontWeight="Bold" Margin="-2,2,0,0"/>
                    <Label x:Name="Label_TexBlock_ChangeSignature_SelectedCert" Content="Selected PFX Cert:" HorizontalAlignment="Left" Margin="-3,64,0,0" VerticalAlignment="Top"/>
                    <Button x:Name="Button_ChangeSignature_SelectFolder" Content="Select Folder" Margin="424,212,0,0" VerticalAlignment="Top" Height="26" HorizontalAlignment="Left" Width="124"/>
                </Grid>
            </TabItem>
            <TabItem x:Name="Tab_AppAttach" Header="AppAttach">
                <Grid Background="#FFE5E5E5" Margin="0,0,0,-183">
                    <Button x:Name="Button_AppAttachConvertToVHD_Convert" Content="Convert to VHD*" Margin="325,124,0,0" VerticalAlignment="Top" Height="26" HorizontalAlignment="Left" Width="128"/>
                    <Label x:Name="Label_AppAttachConvertToVHD_FileOrFolder" Content="What to convert to VHD" HorizontalAlignment="Left" VerticalAlignment="Top" FontWeight="Bold" Margin="-2,3,0,0"/>
                    <TextBlock x:Name="TexBlock_AppAttachConvertToVHD_FileOrFolder" HorizontalAlignment="Left" Margin="3,29,0,0" TextWrapping="Wrap" VerticalAlignment="Top" Height="37" Width="757"><Run Text="This will let you select one single MSIX file or a folder containing MSIX Files"/><Run Text="."/><Run Text=" "/><Run Text="If you select a folder, the tool will also look in it's subfolders for MSIX Files."/></TextBlock>
                    <Button x:Name="Button_AppAttachConvertToVHD_SingelFile_Browse" Content="Select single MSIX" Margin="233,92,0,0" VerticalAlignment="Top" Height="26" HorizontalAlignment="Left" Width="113"/>
                    <Label x:Name="Label_AppAttachConvertToVHD_SingelFile_MSIX" Content="File or Folder:" HorizontalAlignment="Left" Margin="-2,63,0,0" VerticalAlignment="Top"/>
                    <TextBox x:Name="TextBox_AppAttachConvertToVHD_SingelFile_SelectedMSIX" Height="20" Margin="104,67,71,0" TextWrapping="Wrap" VerticalAlignment="Top" MaxHeight="20"/>
                    <Button x:Name="Button_AppAttachConvertToVHD_SelectFolder" Content="Select Folder" Margin="424,92,0,0" VerticalAlignment="Top" Height="26" HorizontalAlignment="Left" Width="124"/>

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

#region Functions

Function ConvertTo-VHD {

    param(
        [Parameter(Mandatory=$true)]
        [String]
        $Path
    )
    #region variables
    $msixmgrURL = "https://aka.ms/msixmgr"
    $TempFolder= "$env:TEMP\_CreatAppAttachMSIX"
    $MsixmgrZip = "$env:TEMP\_msixmgr.zip"
    $MsixmgrFolder = "$TempFolder\_msixmgr"
    $Extractfolder = "$env:SystemDrive\TempExtractAppAttach"
    If([Environment]::Is64BitOperatingSystem){
        $MsixmgrExe = "$MsixmgrFolder\x64\msixmgr.exe"
    }
    else{
        $MsixmgrExe = "$MsixmgrFolder\x86\msixmgr.exe"

    }

    #endregion

    $Scriptblock1Stage = {
#MSIX app attach staging sample


#region variables

$ScriptParentPath = $MyInvocation.MyCommand.Path -replace $myInvocation.MyCommand.Name,""

$vhdSrc = (get-ChildItem -Path $ScriptParentPath -Filter *.vhd)[0].FullName

Write-Host "VHD $vhdSrc found"


$JSONFileName = "AppAttachInfo.json"

$msixJunction = "$Env:Temp\AppAttach\"

#endregion

#region mountvhd

try{
    $VHDFolder = (Get-Item $vhdSrc).DirectoryName
    $JsonFile = "$VHDFolder\$JSONFileName"

    If(Test-Path -Path $JsonFile){
        $Json = get-content $JsonFile |ConvertFrom-Json

        $MountedVHD = Mount-DiskImage -ImagePath $vhdSrc -NoDriveLetter -Access ReadOnly 

        $msixDest = $Json.PartitionPath
        $packageName = $Json.PackageName

        $parentFolder = $Json.ParentFolder
        $parentFolder = "\" + $parentFolder + "\"
    }
    else{
        Write-Host "$JsonFile not found"
    }
}
catch
{
    Write-Host ("Mounting of " + $vhdSrc + " has failed!") -BackgroundColor Red
}

#endregion

#region makelink

if (!(Test-Path $msixJunction)){
    md $msixJunction
}

$msixJunction = $msixJunction + $packageName

cmd.exe /c mklink /j $msixJunction $msixDest

#endregion

#region stage

[Windows.Management.Deployment.PackageManager,Windows.Management.Deployment,ContentType=WindowsRuntime]| Out-Null

Add-Type -AssemblyName System.Runtime.WindowsRuntime

$asTask = ([System.WindowsRuntimeSystemExtensions].GetMethods() | Where {
$_.ToString() -eq 'System.Threading.Tasks.Task`1[TResult] AsTask[TResult,TProgress](Windows.Foundation.IAsyncOperationWithProgress`2[TResult,TProgress])'
})[0]

$asTaskAsyncOperation =
$asTask.MakeGenericMethod([Windows.Management.Deployment.DeploymentResult],
[Windows.Management.Deployment.DeploymentProgress])

$packageManager = [Windows.Management.Deployment.PackageManager]::new()

$path = $msixJunction + $parentFolder + $packageName # needed if we do the pbisigned.vhd

$path = ([System.Uri]$path).AbsoluteUri

$asyncOperation = $packageManager.StagePackageAsync("$path", $null, "StageInPlace")

$task = $asTaskAsyncOperation.Invoke($null, @($asyncOperation))

Do{
    Start-Sleep -Milliseconds  100
}Until(($task.IsCompleted -eq $true) -or ($task.IsFaulted -eq $true))

$task.Exception
$task.Status

#endregion




}

    $Scriptblock2Register = {
#MSIX app attach registration sample

#region variables

#region variables

$ScriptParentPath = $MyInvocation.MyCommand.Path -replace $myInvocation.MyCommand.Name,""

$vhdSrc = (get-ChildItem -Path $ScriptParentPath -Filter *.vhd)[0].FullName

Write-Host "VHD $vhdSrc found"


$JSONFileName = "AppAttachInfo.json"
$VHDFolder = (Get-Item $vhdSrc).DirectoryName
$JsonFile = "$VHDFolder\$JSONFileName"


    If(Test-Path -Path $JsonFile){
        $Json = get-content $JsonFile |ConvertFrom-Json
        $packageName = $Json.PackageName
    }
    else{
        Write-Host "$JsonFile not found"
    }

#endregion

#region register

    $path = "C:\Program Files\WindowsApps\" + $packageName + "\AppxManifest.xml"
    Add-AppxPackage -Path $path -DisableDevelopmentMode -Register


#endregion


    }

    $Scriptblock3Deregister = {
#MSIX app attach deregistration sample



#region variables
$ScriptParentPath = $MyInvocation.MyCommand.Path -replace $myInvocation.MyCommand.Name,""

$vhdSrc = (get-ChildItem -Path $ScriptParentPath -Filter *.vhd)[0].FullName

Write-Host "VHD $vhdSrc found"


$JSONFileName = "AppAttachInfo.json"
$VHDFolder = (Get-Item $vhdSrc).DirectoryName
$JsonFile = "$VHDFolder\$JSONFileName"

If(Test-Path -Path $JsonFile){
    $Json = get-content $JsonFile |ConvertFrom-Json
    $packageName = $Json.PackageName
}
else{
    Write-Host "$JsonFile not found"
}


#endregion

#region deregister

Remove-AppxPackage -PreserveRoamableApplicationData $packageName

#endregion

    }

    $Scriptblock4Destage = {
#MSIX app attach de staging sample


#region variables

$ScriptParentPath = $MyInvocation.MyCommand.Path -replace $myInvocation.MyCommand.Name,""

$vhdSrc = (get-ChildItem -Path $ScriptParentPath -Filter *.vhd)[0].FullName

Write-Host "VHD $vhdSrc found"


$JSONFileName = "AppAttachInfo.json"
$VHDFolder = (Get-Item $vhdSrc).DirectoryName
$JsonFile = "$VHDFolder\$JSONFileName"

$msixJunction = "$Env:Temp\AppAttach\"

#endregion



If(Test-Path -Path $JsonFile){
    $Json = get-content $JsonFile |ConvertFrom-Json
    $packageName = $Json.PackageName
    $PackageMountPoint = "$msixJunction\$packageName"
}
else{
    Write-Host "$JsonFile not found"
}


#region deregister

Remove-AppxPackage -AllUsers -Package $packageName

#cd $msixJunction

Remove-Item -Path $PackageMountPoint -Recurse -Force

# rmdir $packageName -Force -Verbose


Dismount-DiskImage -ImagePath $vhdSrc

#endregion

    }

    #Check if user is Admin
    If ([bool](([System.Security.Principal.WindowsIdentity]::GetCurrent()).groups -match "S-1-5-32-544")) {
        #If he is an Admin

        $VBS = new-object -comobject wscript.shell

        $CleandUpPath = $Path.replace('"',"")


        If(!-(Test-Path $TempFolder)){
            New-Item $TempFolder -ItemType directory
        }

        IF(Test-Path -Path $CleandUpPath){


            If((get-item -Path $CleandUpPath).Extension.ToUpper() -eq ".MSIX"){
                $FoundPackages = get-item -Path $CleandUpPath
            }
            else{
                $FoundPackages = (Get-ChildItem -Path $CleandUpPath -Recurse -Filter "*.msix" ) | Sort-Object
            }


            If($FoundPackages.count -ne 1){
                $WPFTextBox_Messages.Text =  ("In the new window, select all the MSIX packages you want to convert")
                $WPFTextBox_Messages.Foreground = "Black"

                $SelctedPackages = $FoundPackages | Select-Object -Property Name,FullName |Out-GridView -Title ("Select MSIX files you want to convert ("+$FoundPackages.count+")" ) -OutputMode Multiple | Select-Object -Property Fullname,Name
            }
            else{
                $SelctedPackages = $FoundPackages
            }
        

            If($SelctedPackages.FullName.count -ge 1){

                If(-!(Test-Path $MsixmgrExe)){
                    Invoke-WebRequest -Uri $msixmgrURL -OutFile $MsixmgrZip -PassThru -UseBasicParsing

                    If((Test-Path $MsixmgrZip)){
                        Expand-Archive -Path $MsixmgrZip -DestinationPath $MsixmgrFolder
                        If(-!(Test-Path $MsixmgrExe)){
                            $ErrorMessage = $_.Exception.Message
                            $WPFTextBox_Messages.Text =  ("I couldn't extract the Msixmgr Tool from $MsixmgrZip" +" / $ErrorMessage")
                            $WPFTextBox_Messages.Foreground = "Red"
                            Return
                        }
                    }
                    else{
                        $ErrorMessage = $_.Exception.Message
                        $WPFTextBox_Messages.Text =  ("I couldn't download the Msixmgr Tool from $msixmgrURL" +" / $ErrorMessage")
                        $WPFTextBox_Messages.Foreground = "Red"
                        Return
                    }
                }

        
                ForEach($MSIX in $SelctedPackages){
                    $Message = ""

                    try{

                        $MSIXCopy = "$TempFolder\" +$MSIX.Name
                        Copy-Item $MSIX.fullname -Destination $MSIXCopy -Force

                        $MSIXObject = Get-Item $MSIXCopy
                        $MSIXName = $MSIXObject.Name.Replace($MSIXObject.Extension,"").Replace(" ","")

                        $OutPutFolder = "$TempFolder\$MSIXName"

                        If(Test-path $OutPutFolder){
                            Remove-Item $OutPutFolder -Force -Recurse
                        }
                        new-item -Path $OutPutFolder -ItemType Directory

                        $ParentFolder = "PartenFolder_" +$MSIXName
                        $VHDFile = "$OutPutFolder\$MSIXName.vhd"

                        $ZipFile = Copy-Item $MSIXCopy -Destination ("$TempFolder\" +$MSIXObject.Name.Replace($MSIXObject.Extension,".ZIP")) -Force -PassThru

                        If(Test-path $Extractfolder){
                            Remove-Item $Extractfolder -Force -Recurse
                        }

                        Expand-Archive -Path $ZipFile -DestinationPath $Extractfolder -Force
                        

                        $NeededSize = (((Get-ChildItem $Extractfolder -Recurse | measure Length -s).sum) * 2)


                        If($NeededSize -lt 6815744){
                            $NeededSize = 6.5MB
                        }
                        else{
                            $NeededSize = ([System.Math]::ceiling($NeededSize/1MB))*1024*1024
                        }

                        $DiskpartFailed = $false
                        $NeededSize = ([math]::round($NeededSize /1mb)) +21

                        $ArgumentCreateVdisk = "create vdisk file=""$VHDFile"" maximum=$NeededSize type=expandable"
                        $ArgumentCreateVolume = "select vdisk file=""$VHDFile""
                        attach vdisk
                        create partition primary
                        format fs=ntfs
                        assign
                        detail"

                        $DiskPartSrciptFileCreateVHD ="$TempFolder\DiskpartCreateVHD.txt"
                        $DiskPartSrciptFileCreateVolume ="$TempFolder\DiskpartCreateVolume.txt"

                        $ArgumentCreateVdisk | Out-File -FilePath $DiskPartSrciptFileCreateVHD -Force -NoNewline -NoClobber -Encoding utf8
                        $ArgumentCreateVolume | Out-File -FilePath $DiskPartSrciptFileCreateVolume -Force -NoNewline -NoClobber -Encoding utf8

                        $process = Start-Process -FilePath diskpart.exe -ArgumentList "/S $DiskPartSrciptFileCreateVHD" -WindowStyle Hidden -Wait -PassThru

                        If($process.ExitCode -ne 0){
                            $WPFTextBox_Messages.Text =  ("I couldn't create vhd $VHDFile")
                            $WPFTextBox_Messages.Foreground = "Red"
                            return
                        }
                        

                        $VolumesBefore = Get-Volume | Where-Object DriveLetter -ne $null

                        $CreatVolumeReturn=(DISKPART /S $DiskPartSrciptFileCreateVolume)
                        $VolumesAfter= Get-Volume | Where-Object DriveLetter -ne $null

                        #Find the new Volume
                        $NewVolume = @()
                        ForEach($VolumeAfter in $VolumesAfter){
                        $TrueCount = 0
                            ForEach($VolumeBefore in $VolumesBefore){
                                If($VolumeBefore.DriveLetter -eq $VolumeAfter.DriveLetter){
                                    $TrueCount = $TrueCount +1
                                }
                            }
                            If($TrueCount -eq 0){
                                $NewVolume += $VolumeAfter   
                            }
                        }


                        If($NewVolume.count -lt 1){
                            $WPFTextBox_Messages.Text =  ("Found no new driveletter mounting the vhd failed")
                            $WPFTextBox_Messages.Foreground = "Red"
                            return
                        }
                        elseIf($NewVolume.count -gt 1){
                            $WPFTextBox_Messages.Text =  ("Found to many new driveletters please try again")
                            $WPFTextBox_Messages.Foreground = "Red"
                            return
                        }                        
                        
                        $PartenFolderPath = ($NewVolume.DriveLetter +":\$ParentFolder").tostring()



                        $Destiantion = $PartenFolderPath

                        $MisxmgrArgument = "-Unpack -packagePath ""$MSIXCopy"" -destination ""$Destiantion"" -applyacls"

                        $Result = Start-Process -FilePath $MsixmgrExe -ArgumentList ($MisxmgrArgument) -Wait -PassThru -WindowStyle Hidden


                        IF($Result.ExitCode -ne 0){
                            $WPFTextBox_Messages.Text =  ("Failed to extract MSIX using: $MsixmgrExe $MisxmgrArgument" )
                            $WPFTextBox_Messages.Foreground = "Red"
                            return
                        }

                        $PackageName = (Get-ChildItem -Path $Destiantion -Directory).Name

                        If($PackageName){

                            $AppAttachInfo = New-Object -TypeName psobject 

                            $AppAttachInfo | Add-Member -MemberType NoteProperty -Name MSIXName -Value $MSIXObject.Name
                            $AppAttachInfo | Add-Member -MemberType NoteProperty -Name PartitionPath -Value $NewVolume.Path
                            $AppAttachInfo | Add-Member -MemberType NoteProperty -Name ParentFolder -Value $ParentFolder
                            $AppAttachInfo | Add-Member -MemberType NoteProperty -Name PackageName -Value $PackageName
                        }
                        else{
                            throw
                        }

                        $AppAttachInfo | ConvertTo-Json | Out-File -FilePath ("$OutPutFolder\AppAttachInfo.json")
                                                

                        $CertContent= (Get-AuthenticodeSignature -FilePath $MSIXCopy ).SignerCertificate
                        $CertToExport = "$OutPutFolder\$MSIXName.cer"
                        If($CertContent){
                            Export-Certificate -Cert $CertContent -FilePath $CertToExport -Force
                        }

                        #Generate Mountig Scripts
                        $Scriptblock1Stage.ToString() | out-file -FilePath "$OutPutFolder\1_Stage.ps1" -Force
                        $Scriptblock2Register.ToString() | out-file -FilePath "$OutPutFolder\2_Register.ps1" -Force
                        $Scriptblock3Deregister.ToString() | out-file -FilePath "$OutPutFolder\3_Deregister.ps1" -Force
                        $Scriptblock4Destage.ToString() | out-file -FilePath "$OutPutFolder\4_Destage.ps1" -Force

                    }catch{
                        $ErrorMessage = $_.Exception.Message

                        $WPFTextBox_Messages.Text =  ("Failed to convert MSIX " +$MSIXName +" / $ErrorMessage")
                        $WPFTextBox_Messages.Foreground = "Red"

                        $VBS.popup(("Failed to convert MSIX " +$MSIXName +"`n`n$ErrorMessage"),0,"Error",16)
                    }
                    Finally{
                        #CleanUp
                        If(Test-Path $DiskPartSrciptFileCreateVHD){
                            remove-item -Path $DiskPartSrciptFileCreateVHD -Force
                        }

                        If(Test-Path $DiskPartSrciptFileCreateVolume){
                            remove-item -Path $DiskPartSrciptFileCreateVolume -Force
                        }

                        Remove-Item $ZipFile -Force -ErrorAction SilentlyContinue
                        Remove-Item $Extractfolder -Force -Recurse -ErrorAction SilentlyContinue
                        Remove-Item -Path $MSIXCopy -Force -ErrorAction SilentlyContinue

                        #Detach VHD
                        $ArgumentDetachVDisk = "select vdisk file=""$VHDFile""
                        DETACH vdisk"

                        $DiskPartSrciptFileDetachVDisk ="$TempFolder\DiskpartDetachVDisk.txt"

                        $ArgumentDetachVDisk | Out-File -FilePath $DiskPartSrciptFileDetachVDisk -Force -NoNewline -NoClobber -Encoding utf8

                        $process = Start-Process -FilePath diskpart.exe -ArgumentList "/S $DiskPartSrciptFileDetachVDisk" -WindowStyle Hidden -Wait -PassThru

                        If(Test-Path $DiskPartSrciptFileDetachVDisk){
                            remove-item -Path $DiskPartSrciptFileDetachVDisk -Force
                        }
                        If($process.ExitCode -ne 0){
                            $WPFTextBox_Messages.Text =  ("Failed to detach $VHDFile")
                            $WPFTextBox_Messages.Foreground = "Red"
                        }
                        
                    }

                }

                $WPFTextBox_Messages.Text =  ("All selected " +$SelctedPackages.count +" Packages are converted.")
                $WPFTextBox_Messages.Foreground = "Black"
                explorer.exe $TempFolder
            }
            else{
                $WPFTextBox_Messages.Text =  ("No MSIX packages selected")
                $WPFTextBox_Messages.Foreground = "Red"
            }
        }
        else{
            $WPFTextBox_Messages.Text =  ("File or Folder not found $CleandUpPath")
            $WPFTextBox_Messages.Foreground = "Red"
        }
    }
    else{
         #If he is NO Admin
        $WPFTextBox_Messages.Text = "You are not an administrator. This Action requieres the tool to be run as an administrator!"
        $WPFTextBox_Messages.Foreground = "Red"
        Return
    }

}

Function Open-Tool{

    param(
        [Parameter(Mandatory=$true)]
        [String]
        $tool
    )
    $SelectedMSIX = $null
    $SelectedMSIX = $WPFListView_MSIXPackages.SelectedItem.Name
    $SelectedMSIXInstallLocation= $WPFListView_MSIXPackages.SelectedItem.InstallLocation

    $SelectedMSIXInstallPublisherHash = $SelectedMSIXInstallLocation.Substring($SelectedMSIXInstallLocation.LastIndexOf("_")+1,$SelectedMSIXInstallLocation.Length -$SelectedMSIXInstallLocation.LastIndexOf("_")-1)
    $UserDataLocation = "$env:LOCALAPPDATA\Packages\$SelectedMSIX"+"_"+$SelectedMSIXInstallPublisherHash


    If($MSIXData){
        IF($SelectedMSIX){
            try{
                $Package = Get-AppxPackage -Name $SelectedMSIX
                $Manifest = Get-AppxPackageManifest -package $Package
                $AppId = $Manifest.package.Applications.Application.Id
                Invoke-CommandInDesktopPackage -PackageFamilyName $Package.PackageFamilyName  -PreventBreakaway -command $tool -AppId $AppId
                $WPFTextBox_Messages.Text =  "Opend $tool inside MSIX $SelectedMSIX"
                $WPFTextBox_Messages.Foreground = "Black"
            }
            catch{
                $ErrorMessage = $_.Exception.Message
                $WPFTextBox_Messages.Text =  ("Couldn't opend $tool inside MSIX $SelectedMSIX / $ErrorMessage")
                $WPFTextBox_Messages.Foreground = "Red"
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

Function Enable-DeveloperMode{

    try{
        $registryPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\AppModelUnlock" 
        $Name1 = "AllowAllTrustedApps"
        $Name2 = "AllowDevelopmentWithoutDevLicense"
        $value1 = "1"
        $value2 = "1"

        New-ItemProperty -Path $registryPath -Name $name1 -Value $value1 -PropertyType DWORD -Force -ErrorAction Stop
        New-ItemProperty -Path $registryPath -Name $name2 -Value $value2 -PropertyType DWORD -Force -ErrorAction Stop

        $WPFTextBox_Messages.Text =  ("Developer Mode succesfully enablead")
        $WPFTextBox_Messages.Foreground = "Green"

    }
    catch{
        $ErrorMessage = $_.Exception.Message
        $WPFTextBox_Messages.Text =  ("Couldn't enable Developer Mode / $ErrorMessage")
        $WPFTextBox_Messages.Foreground = "Red"
    }
}

Function Enable-Sideloading{

    try{
        $registryPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\AppModelUnlock" 
        $Name1 = "AllowAllTrustedApps"
        $Name2 = "AllowDevelopmentWithoutDevLicense"
        $value1 = "1"
        $value2 = "0"

        New-ItemProperty -Path $registryPath -Name $name1 -Value $value1 -PropertyType DWORD -Force -ErrorAction Stop
        New-ItemProperty -Path $registryPath -Name $name2 -Value $value2 -PropertyType DWORD -Force -ErrorAction Stop

        $WPFTextBox_Messages.Text =  ("Sideloading succesfully enablead")
        $WPFTextBox_Messages.Foreground = "Green"

    }
    catch{
        $ErrorMessage = $_.Exception.Message
        $WPFTextBox_Messages.Text =  ("Couldn't enable Sideloading / $ErrorMessage")
        $WPFTextBox_Messages.Foreground = "Red"
    }
}

Function Disable-Sideloading{

    try{
        $registryPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\AppModelUnlock" 
        $Name1 = "AllowAllTrustedApps"
        $Name2 = "AllowDevelopmentWithoutDevLicense"
        $value1 = "0"

        New-ItemProperty -Path $registryPath -Name $name1 -Value $value1 -PropertyType DWORD -Force -ErrorAction Stop
        New-ItemProperty -Path $registryPath -Name $name2 -Value $value1 -PropertyType DWORD -Force -ErrorAction Stop

        $WPFTextBox_Messages.Text =  ("Sideloading succesfully disabled")
        $WPFTextBox_Messages.Foreground = "Green"

    }
    catch{
        $ErrorMessage = $_.Exception.Message
        $WPFTextBox_Messages.Text =  ("Couldn't disable Sideloading / $ErrorMessage")
        $WPFTextBox_Messages.Foreground = "Red"
    }
}

function Check-SideloadingStaus{

    $registryPath = "HKLM:\SOFTWARE\Microsoft\Windows\CurrentVersion\AppModelUnlock" 
    $Name1 = "AllowAllTrustedApps"
    $Name2 = "AllowDevelopmentWithoutDevLicense"

    If((Get-ItemProperty -Path $registryPath -ErrorAction SilentlyContinue).$Name1 -eq 1){
        If((Get-ItemProperty -Path $registryPath -ErrorAction SilentlyContinue).$Name2 -eq 1){
            "Developer Mode Enabled"
        }
        else{
            "Enabled"
        }
    }
    else{
        "Disabled"
    }

}

Function Change-Signature {

    param(
        [Parameter(Mandatory=$true)]
        [String]
        $Path
    )

    $VBS = new-object -comobject wscript.shell


    $CleandUpPath = $Path.replace('"',"")

    $pfxPath = ($WPFTextBox_ChangeSignature_SelectedCert.Text).replace('"',"")
    $Password = $WPFPasswordBox_ChangeSignatureCertPassword.Password

    $TimeStampServer = $WPFTextBox_ChangeSignature_Timeserver.Text

        
    If(($pfxPath) -and ($Password)){
        #Get the SDK Tools
        $MSIXPackagingToolInstallLocation =  (Get-AppxPackage -Name "Microsoft.MsixPackagingTool*").InstallLocation
        If($MSIXPackagingToolInstallLocation){

                    $MsixPackagingToolSDK = "$env:TEMP\MsixPackagingToolSDK"
                    Copy-Item -Path "$MSIXPackagingToolInstallLocation\SDK" -Destination $MsixPackagingToolSDK  -Recurse -Force

                    $makeappxPath = "$MsixPackagingToolSDK\makeappx.exe"
                    $SignTool = "$MsixPackagingToolSDK\signtool.exe"

                    $NewMSIXFolder = "$env:TEMP\_NewMSIXFiles\"

                    IF(Test-Path -Path $CleandUpPath){
                        If((get-item -Path $CleandUpPath).Extension.ToUpper() -eq ".MSIX"){
                            $SelctedPackages = get-item -Path $CleandUpPath
                        }
                        else{
                            $SelctedPackages = (Get-ChildItem -Path $CleandUpPath -Recurse -Filter "*.msix" ) | Sort-Object
                        }

                        ForEach($package in $SelctedPackages){

                            try{

                                $SelectedMSIX = get-item -path $package.fullname

                                $ExtractFolderName = $SelectedMSIX.Name.Replace($SelectedMSIX.Extension,"")

                                $ExtractPath = "$env:TEMP\$ExtractFolderName"

                                If(Test-Path -Path $ExtractPath){
                                    Remove-Item -Force -Path $ExtractPath -Recurse
                                }

                                $UnpackParmas ="unpack -v -d ""$ExtractPath"" -p ""$SelectedMSIX"" -o"


                                #Extract MSIX
                                $process = Start-Process -FilePath $makeappxPath -ArgumentList $UnpackParmas -Wait -WindowStyle Hidden -PassThru

                                If($process.exitcode -ne 0){
                                    write-host $makeappxPath
                                    write-host $UnpackParmas
                                    throw
                                }                                

                                $ManifestXMLPath = "$ExtractPath\AppxManifest.xml"

                                $ManifestContent = [XML](Get-Content -Path $ManifestXMLPath)          

                                #Get infos from cert
                                $cert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2
                                $cert.Import($pfxPath,$Password,'DefaultKeySet')

                                $NewPublisher = $cert.Subject
                                $OldPublisher = $ManifestContent.Package.Identity.Publisher

                                If($OldPublisher -ne $NewPublisher){
                                    $ManifestContent.Package.Identity.Publisher = $NewPublisher
                                }

                                #Check if Modification Package
                                $ModificationPackagePublisher = $ManifestContent.Package.Dependencies.MainPackageDependency.Publisher
                                If($OldPublisher -eq $ModificationPackagePublisher){
                                    $Answer =$VBS.Popup(($package.Name +" is a modifiaction Package.`nThe Publisher of this Package and of the MainPackageDependency are identical.`n`nDo you want me to also update the MainPackageDependency publisher?"),0,"Modification Package",36)
                                    If($Answer -eq  6){
                                        $ManifestContent.Package.Dependencies.MainPackageDependency.Publisher = $NewPublisher
                                    }
                                }

                                #Save NewXML
                                $ManifestContent.Save($ManifestXMLPath)

                                If((Test-Path $NewMSIXFolder) -eq $false){
                                    md $NewMSIXFolder
                                }

                                $NewMSIXPath = "$NewMSIXFolder\$ExtractFolderName.msix"

                                $MakeAppxParmas ="pack -d ""$ExtractPath"" -p ""$NewMSIXPath"" -o"

                                $process = Start-Process -FilePath $makeappxPath -ArgumentList $MakeAppxParmas -PassThru -Wait -WindowStyle Hidden 

                                If($process.exitcode -ne 0){                                    
                                    write-host $makeappxPath
                                    write-host $MakeAppxParmas
                                    throw
                                }

                                Remove-Item -Force -Path $ExtractPath -Recurse

                                If($pfxPath -and $Password){
                                    If($TimeStampServer){
                                        $SignArguments = "sign -f ""$pfxPath"" -p $Password -t ""$TimeStampServer"" -fd SHA256 -v ""$NewMSIXPath"""
                                    }
                                    else{
                                        $SignArguments = "sign -f ""$pfxPath"" -p $Password -fd SHA256 -v ""$NewMSIXPath"""
                                    }

                                    $process = Start-Process -FilePath $SignTool -ArgumentList ($SignArguments) -PassThru -Wait -WindowStyle Hidden

                                    If($process.exitcode -ne 0){
                                        write-host $SignTool
                                        write-host $SignArguments
                                        throw
                                    }
                                    #Check if the Name contians __ folowed by 13 charachters because then that would be the oldpublisherID
                                    If($ExtractFolderName.Length - $ExtractFolderName.LastIndexOf("__") -2 -eq 13){
                                        #Install to get the Publisher ID and rename the file. After that uninstall it again

                                        try{
                                            $OldPublisherID = $ExtractFolderName.Substring($ExtractFolderName.LastIndexOf("__")+2, $ExtractFolderName.Length - $ExtractFolderName.LastIndexOf("__")-2)

                                            $Package = get-item $NewMSIXPath
                                            
                                            $AppxPackagesBefore = get-AppxPackage
                                            Add-AppxPackage -Path $package.Fullname -ErrorAction Stop
                                            $AppxPackagesAfter = get-AppxPackage
                                            $Difference = Compare-Object -ReferenceObject $AppxPackagesBefore -DifferenceObject $AppxPackagesAfter
            
                                            If($Difference.InputObject.count -gt 0){

                                                $NewPublisherId = $Difference.InputObject.PublisherId

                                                Rename-Item -path $NewMSIXPath -NewName ($ExtractFolderName.Replace($OldPublisherID,$NewPublisherId)+".msix")

                                                try{
                                                    $Difference.InputObject | Remove-AppxPackage -ErrorAction Stop

                                                    $WPFTextBox_Messages.Text =  ("Removed MSIX " +$package.Name)
                                                    $WPFTextBox_Messages.Foreground = "Black"

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

                                            If($ModificationPackagePublisher.Length -gt 2){

                                                $WPFTextBox_Messages.Text =  ("Failed to add MSIX " +$package.Name +" it seams to be a Modification Package / $ErrorMessage")
                                                $WPFTextBox_Messages.Foreground = "Red"
                                                $VBS.popup(($package.Name +"`nIs a Modification Package. I couldn't automatically change the publisher hash in the Filename.`nYou need to change it manually."),0,"Modification Package",48)
                                            }
                                            else{

                                                $ErrorMessage = $_.Exception.Message
                                                $WPFTextBox_Messages.Text =  ("Failed to add MSIX " +$package.Name +" / $ErrorMessage")
                                                $WPFTextBox_Messages.Foreground = "Red"

                                                $VBS.popup(("Failed to add MSIX " +$package.Name +"`n`n$ErrorMessage"),0,"Error",16)
                                            }
                                        }
                                    }
                                }
                            }catch{
                                $ErrorMessage = $_.Exception.Message

                                $WPFTextBox_Messages.Text =  ("Failed to changed the signature for package " +$package.Name +" / $ErrorMessage")
                                $WPFTextBox_Messages.Foreground = "Red"
                                return
                            }
                        }

                        $WPFTextBox_Messages.Text =  ("Change the signature for all selected " +$SelctedPackages.count +" packages.")
                        $WPFTextBox_Messages.Foreground = "Black"

                        explorer.exe $NewMSIXFolder
                    }
                    else{
                        $WPFTextBox_Messages.Text =  ("File or Folder not found $CleandUpPath")
                        $WPFTextBox_Messages.Foreground = "Red"
                    }

                    Remove-Item -Force -Path $MsixPackagingToolSDK -Recurse
        }
        else{
            $WPFTextBox_Messages.Text =  ("Failed to find the MSXI Packaging Tool. For this functionality it is required. Pleas install it from the Windows Store.")
            $WPFTextBox_Messages.Foreground = "Red"
        }
    }else{
        $WPFTextBox_Messages.Text =  ("No pfx cert and password, therefore I don’t need to modify anything.")
        $WPFTextBox_Messages.Foreground = "Red"
    }
}

Function Stop-Services {

    #List Of Services to stopp and disable
    $StoppedServices = "none"
    $FailedServices = "none"
    $NotFoundServices = "none"
    $Services = @("PLRestartMgrService","DPS","CscService","WSearch","wuauserv","CcmExec")


    If($UserIsAdmin -eq $true){
        #Disable Services

        ForEach ($Service in $Services){
            $IsAvailable = $null
            $IsAvailable = Get-service $Service -ErrorAction SilentlyContinue
            If($IsAvailable){
                try{
                    If($IsAvailable.Status -eq "Running"){
                        Stop-Service -Name $Service -Force -ErrorAction Stop
                    }
                    Set-Service –Name $Service –StartupType "Disabled" -ErrorAction Stop
            
                    If($StoppedServices -ne "none"){
                        $StoppedServices += (", " +$IsAvailable.DisplayName)
                    }
                    else{
                        $StoppedServices = $IsAvailable.DisplayName.ToString()
                    }
                }
                catch{
                    If($FailedServices -ne "none"){
                        $FailedServices += (", " +$IsAvailable.DisplayName)
                    }
                    else{
                        $FailedServices = $IsAvailable.DisplayName.ToString()
                    }
                }
            }
            else{
                If($NotFoundServices -ne "none"){
                    $NotFoundServices += (", " +$Service)
                }
                else{
                    $NotFoundServices = $Service
                }
            }
        }

        $Message = "Stopped those Services: $StoppedServices / Couldn't find those: $NotFoundServices / Failed to stop those: $FailedServices"
        $WPFTextBox_Messages.Text =  $Message

        If($FailedServices -ne "none"){
            $WPFTextBox_Messages.Foreground = "Red"
        }
        else{
            $WPFTextBox_Messages.Foreground = "Black"
        }
    }
    else{
        $WPFTextBox_Messages.Text =  "You are not an administrator and this functionality requires admin rights.”
        $WPFTextBox_Messages.Foreground = "Red"
    }
}

Function Edit-Manifest {

    param(
        [Parameter(Mandatory=$true)]
        [String]
        $Path
    )

    $VBS = new-object -comobject wscript.shell


    $RegExForBuild = "(0|[1-9][0-9]{0,3}|[1-5][0-9]{4}|6[0-4][0-9]{3}|65[0-4][0-9]{2}|655[0-2][0-9]|6553[0-5])(\.(0|[1-9][0-9]{0,3}|[1-5][0-9]{4}|6[0-4][0-9]{3}|65[0-4][0-9]{2}|655[0-2][0-9]|6553[0-5])){3}"
    $CleandUpPath = $Path.replace('"',"")

    $pfxPath = $WPFTextBox_EditManifest_SelectedCert.Text
    $Password = $WPFPasswordBox_EditManifest_CertPassword.Password

    $TimeStampServer = $WPFTextBox_Timeserver.Text

    $NewMaxVersionTested = $WPFTextBox_MaxVersionTested.Text
    $NewMinVersionTested = $WPFTextBox_MinVersionTested.Text

        
    If(($NewMaxVersionTested) -or ($NewMinVersionTested)){
        #Check if Versionnumbers are Build numbers
        If(($NewMinVersionTested -match $RegExForBuild) -eq $false){
            $WPFTextBox_Messages.Text =  ("$NewMinVersionTested is not a regular Buildnumber like 1.0.0.0")
            $WPFTextBox_Messages.Foreground = "Red"
            Return
        }

        If(($NewMaxVersionTested -match $RegExForBuild) -eq $false){
            $WPFTextBox_Messages.Text =  ("$NewMaxVersionTested is not a regular Buildnumber like 1.0.0.0")
            $WPFTextBox_Messages.Foreground = "Red"
            Return
        }

        #Get the SDK Tools
        $MSIXPackagingToolInstallLocation =  (Get-AppxPackage -Name "Microsoft.MsixPackagingTool*").InstallLocation

        If($MSIXPackagingToolInstallLocation){

                    $MsixPackagingToolSDK = "$env:TEMP\MsixPackagingToolSDK"
                    Copy-Item -Path "$MSIXPackagingToolInstallLocation\SDK" -Destination $MsixPackagingToolSDK  -Recurse -Force

                    $makeappxPath = "$MsixPackagingToolSDK\makeappx.exe"
                    $SignTool = "$MsixPackagingToolSDK\signtool.exe"

                    $NewMSIXFolder = "$env:TEMP\_NewMSIXFiles\"

                    IF(Test-Path -Path $CleandUpPath){
                        If((get-item -Path $CleandUpPath).Extension.ToUpper() -eq ".MSIX"){
                            $SelctedPackages = get-item -Path $CleandUpPath
                        }
                        else{
                            $SelctedPackages = (Get-ChildItem -Path $CleandUpPath -Recurse -Filter "*.msix" ) | Sort-Object
                        }

                        ForEach($package in $SelctedPackages){

                            try{

                                $SelectedMSIX = get-item -path $package.fullname

                                $ExtractFolderName = $SelectedMSIX.Name.Replace($SelectedMSIX.Extension,"")

                                $ExtractPath = "$env:TEMP\$ExtractFolderName"

                                If(Test-Path -Path $ExtractPath){
                                    Remove-Item -Force -Path $ExtractPath -Recurse
                                }

                                $UnpackParmas ="unpack -v -d ""$ExtractPath"" -p ""$SelectedMSIX"" -o"


                                #Extract MSIX
                                $process = Start-Process -FilePath $makeappxPath -ArgumentList $UnpackParmas -Wait -WindowStyle Hidden -PassThru

                                If($process.exitcode -ne 0){
                                    write-host $makeappxPath
                                    write-host $UnpackParmas
                                    throw
                                }                                

                                $ManifestXMLPath = "$ExtractPath\AppxManifest.xml"

                                $ManifestContent = [XML](Get-Content -Path $ManifestXMLPath)
                                                                
                                $MSIXVersionOld = $ManifestContent.Package.Identity.Version


                                If($WPFRadioButton_EditManifest_IncreaseVersion_Yes.IsChecked){
                                    $MSIXVersionNew = $MSIXVersionOld.split(".")[0] +"." + $MSIXVersionOld.split(".")[1] +"." + $MSIXVersionOld.split(".")[2] +"." +([INT]$MSIXVersionOld.split(".")[3] +1)
                                    #Set Increased Version
                                    $ManifestContent.Package.Identity.Version = $MSIXVersionNew
                                    $NewMSIXName = $ExtractFolderName.Replace($MSIXVersionOld,$MSIXVersionNew)
                                }
                                else{
                                    $NewMSIXName = $ExtractFolderName
                                }

                                                                   
                                #Change MaxVersionTested and Minversion
                                If($NewMaxVersionTested){                                  
                                    $ManifestContent.Package.Dependencies.TargetDeviceFamily.MaxVersionTested = $NewMaxVersionTested    
                                }

                                If($NewMinVersionTested){
                                    $ManifestContent.Package.Dependencies.TargetDeviceFamily.MinVersion = $NewMinVersionTested
                                }

                                If(($pfxPath) -and ($Password)){
                                    $cert = New-Object System.Security.Cryptography.X509Certificates.X509Certificate2
                                    $cert.Import($pfxPath,$Password,'DefaultKeySet')

                                    $NewPublisher = $cert.Subject
                                    $OldPublisher = $ManifestContent.Package.Identity.Publisher

                                    If($OldPublisher -ne $NewPublisher){

                                        $Answer = $VBS.popup(($SelectedMSIX.Name +"`n`n" +"You are not using the same certificate that the one the original package was signed with.`n`nIf you use this certificate the package will be treated as a new package. This means that Updates and Modification Packages will no longer work.`n`nIt’s heavily recommended to use the original certificate and now click on No. `n`nDo you want to use this new certificate anyways and ignore this warning?"),0,"Publisher certificate is different",36)

                                        #7 means the user pressed No, so the modification failes
                                        If($Answer -eq 7){                                        
                                            throw
                                        }

                                        $ManifestContent.Package.Identity.Publisher = $NewPublisher
                                    }
                                }

                                #Save NewXML
                                $ManifestContent.Save($ManifestXMLPath)

                                If((Test-Path $NewMSIXFolder) -eq $false){
                                    md $NewMSIXFolder
                                }

                                $NewMSIXPath = "$NewMSIXFolder\$NewMSIXName.msix"

                                $MakeAppxParmas ="pack -d ""$ExtractPath"" -p ""$NewMSIXPath"" -o"

                                $process = Start-Process -FilePath $makeappxPath -ArgumentList $MakeAppxParmas -PassThru -Wait -WindowStyle Hidden 

                                If($process.exitcode -ne 0){                                    
                                    write-host $makeappxPath
                                    write-host $MakeAppxParmas
                                    throw
                                }

                               # Remove-Item -Force -Path $ExtractPath -Recurse

                                If($pfxPath -and $Password){
                                    If($TimeStampServer){
                                        $SignArguments = "sign -f ""$pfxPath"" -p $Password -t ""$TimeStampServer"" -fd SHA256 -v ""$NewMSIXPath"""
                                    }
                                    else{
                                        $SignArguments = "sign -f ""$pfxPath"" -p $Password -fd SHA256 -v ""$NewMSIXPath"""
                                    }

                                    $process = Start-Process -FilePath $SignTool -ArgumentList ($SignArguments) -PassThru -Wait -WindowStyle Hidden

                                    If($process.exitcode -ne 0){
                                        write-host $SignTool
                                        write-host $SignArguments
                                        throw
                                    }
                                }
                            }catch{
                                $ErrorMessage = $_.Exception.Message

                                $WPFTextBox_Messages.Text =  ("Failed to edit manifest for package " +$package.Name +" / $ErrorMessage")
                                $WPFTextBox_Messages.Foreground = "Red"
                                return
                            }

                        }

                        $WPFTextBox_Messages.Text =  ("All selected " +$SelctedPackages.count +" Packages are edited.")
                        $WPFTextBox_Messages.Foreground = "Black"

                        explorer.exe $NewMSIXFolder
                    }
                    else{
                        $WPFTextBox_Messages.Text =  ("File or Folder not found $CleandUpPath")
                        $WPFTextBox_Messages.Foreground = "Red"
                    }

                    Remove-Item -Force -Path $MsixPackagingToolSDK -Recurse
        }
        else{
            $WPFTextBox_Messages.Text =  ("Failed to find the MSXI Packaging Tool. For this functionality it is required. Pleas install it from the Windows Store.")
            $WPFTextBox_Messages.Foreground = "Red"
        }
    }else{
        $WPFTextBox_Messages.Text =  ("No new MinVersionTested or MaxVersionTested specified, therefore I don’t need to modify anything.")
        $WPFTextBox_Messages.Foreground = "Orange"

    }

}

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

Function Start-App{
    $SelectedMSIX = $null
    $SelectedMSIX = $WPFListView_MSIXPackages.SelectedItem.Name
    $SelectedMSIXInstallLocation= $WPFListView_MSIXPackages.SelectedItem.InstallLocation
    $SelectedMSIXInstallLocation= $WPFListView_MSIXPackages.SelectedItem.InstallLocation
    $SelectedMSIXManifest = "$SelectedMSIXInstallLocation\AppxManifest.xml"
    
    If($MSIXData){
        IF($SelectedMSIX){
            $WPFTextBox_Messages.Text =  "The selected MSIX is $SelectedMSIX"
            $WPFTextBox_Messages.Foreground = "Black"

             [XML]$SelectedMSIXManifestContent = Get-Content $SelectedMSIXManifest
            $ApplicationIDs = ($SelectedMSIXManifestContent.Package.Applications.Application).id
  
       
            If($ApplicationIDs.count -le 1){
                try{
                    ForEach($ApplicationID in $ApplicationIDs){
                        explorer.exe shell:AppsFolder\$(get-appxpackage -name $SelectedMSIX | select -expandproperty PackageFamilyName)!$ApplicationID
                    }
                        $WPFTextBox_Messages.Text =  "Startet the Apps from the MSIX $SelectedMSIX"
                        $WPFTextBox_Messages.Foreground = "Black"
                    }
                catch{
                    $ErrorMessage = $_.Exception.Message

                    $WPFTextBox_Messages.Text ="Failed to Start the Apps from  $SelectedMSIX / $ErrorMessage"
                    $WPFTextBox_Messages.Foreground = "RED"
                }
            }
            else{
                $WPFTextBox_Messages.Text =  "Found no Application ID's in the $SelectedMSIXManifest"
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
                Use-Filter
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

Function Open-Manifest{
    $SelectedMSIX = $null
    $SelectedMSIX = $WPFListView_MSIXPackages.SelectedItem.Name
    $SelectedMSIXInstallLocation= $WPFListView_MSIXPackages.SelectedItem.InstallLocation

    If($MSIXData){
        IF($SelectedMSIX){
            $WPFTextBox_Messages.Text =  "The selected MSIX is $SelectedMSIX"
            $WPFTextBox_Messages.Foreground = "Black"
            explorer.exe "$SelectedMSIXInstallLocation\AppxManifest.xml"


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
        $InstalledApps = @()


    If($WPFCheckBox_EnterpriseSigned.IsChecked){
        $InstalledApps += Get-AppxPackage | Where-object SignatureKind -EQ "Developer"  |Select-Object -Property Name,Version,Publisher,InstallLocation,Dependencies,PackageFullName
        $InstalledApps += Get-AppxPackage | Where-object SignatureKind -EQ "Enterprise"  |Select-Object -Property Name,Version,Publisher,InstallLocation,Dependencies,PackageFullName  

    }
    else{

        $InstalledApps = Get-AppxPackage | Select-Object -Property Name,Version,Publisher,InstallLocation,Dependencies,PackageFullName 

    }

    ForEach($InstalledApp in $InstalledApps){
        $Dependencies = ""
        $DependenciesCount = 0
        $Name = $InstalledApp.Name
        $Publisher = $InstalledApp.Publisher
        $Version = $InstalledApp.Version
        $InstallLocation = $InstalledApp.InstallLocation
        ForEach($Dependencie in $InstalledApp.Dependencies){
            $DependenciesCount = $DependenciesCount +1
            If($DependenciesCount -eq 1){
                $Dependencies  = $Dependencie.ToString()
            }else{
                $Dependencies  = $Dependencies +" / " +$Dependencie.ToString()

            }
        }

        $PackageType = 'Appx'
        $mani =  Get-AppxPackageManifest -Package $InstalledApp.PackageFullName
        $capabilitiesArr = $mani.GetElementsByTagName('Capabilities')
        ForEach($capabilities in $capabilitiesArr)
        {
            $rescapCapabilityArr = $capabilities.GetElementsByTagName('rescap:Capability')
            foreach ($rescapCapability in $rescapCapabilityArr)
            {
                if ($rescapCapability.Name -eq 'runFullTrust')
                {
                    $PackageType = 'MSIX (Full Trust)'
                }
            }
            $capabilityArr = $capabilities.GetElementsByTagName('Capability')
            foreach ($capability in $capabilityArr)
            {
                #only possible in manually generated manifests
                if ($capability.Name -eq 'runFullTrust')
                {
                    $PackageType = 'MSIX (Full Trust)'
                }
            }
        }

        #$InstallDate = [Datetime]::ParseExact($InstalledSoftware.InstallDate,(Get-culture).DateTimeFormat.ShortDatePattern +" " +(Get-culture).DateTimeFormat.LongTimePattern,$null)

        $SoftwareDetail = New-Object PSObject
        $SoftwareDetail | Add-Member -Name "Name" -MemberType NoteProperty -Value $Name
        $SoftwareDetail | Add-Member -Name "Version" -MemberType NoteProperty -Value $Version
        $SoftwareDetail | Add-Member -Name "Publisher" -MemberType NoteProperty -Value $Publisher
        $SoftwareDetail | Add-Member -Name "InstallLocation" -MemberType NoteProperty -Value $InstallLocation
        $SoftwareDetail | Add-Member -Name "PackageType" -MemberType NoteProperty -Value $PackageType
        $SoftwareDetail | Add-Member -Name "Dependencies" -MemberType NoteProperty -Value $Dependencies

        $Script:MSIXData += $SoftwareDetail

        
    }


    If ($Script:MSIXData.count -gt 1){
        $sortedlist = ($Script:MSIXData | Sort-Object -Property Name)
    }
    else{
        $sortedlist = @($Script:MSIXData,"")
    }


    $WPFListView_MSIXPackages.ItemsSource = $sortedlist

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

Function Import-CertToTrustedPeopleFromMSIX{

    #Check if user is Admin
    If ([bool](([System.Security.Principal.WindowsIdentity]::GetCurrent()).groups -match "S-1-5-32-544")) {
        #If he is an Admin



        $MSIXToInstallCertFrom = $WPFTextBox_SelectedCertFromMSIX.Text
        If(-!$MSIXToInstallCertFrom){
            $WPFTextBox_Messages.Text ="No MSIX provided"
            $WPFTextBox_Messages.Foreground = "RED"
            Return
        }
        else{
            If(Test-Path -Path $MSIXToInstallCertFrom){

                try{

                    $CertContent= (Get-AuthenticodeSignature -FilePath $MSIXToInstallCertFrom).SignerCertificate
                    $CertToInstall = "$env:temp\temp.cer"
                    Export-Certificate -Cert $CertContent -FilePath $CertToInstall -Force
                    Import-Certificate -FilePath $CertToInstall -CertStoreLocation cert:\LocalMachine\TrustedPeople
                    Remove-Item -Path $CertToInstall -Force

                    $WPFTextBox_Messages.Text ="Succesfully installed Cert to LocalMachine\TrustedPeople from $MSIXToInstallCertFrom"
                    $WPFTextBox_Messages.Foreground = "Black"

                }
                catch{
                    $ErrorMessage = $_.Exception.Message

                    $WPFTextBox_Messages.Text ="Failed to install the cert File / $ErrorMessage"
                    $WPFTextBox_Messages.Foreground = "RED"
                }

            }
            else{
                $WPFTextBox_Messages.Text ="The Path $MSIXToInstallCertFrom is not valid"
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

Function Select-File {
    param(
        [Parameter(Mandatory=$true)]
        [String]
        $FileType
    )

    #Select File
    [System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null    
    $OpenFileDialog = New-Object System.Windows.Forms.OpenFileDialog
    If($FileType.IndexOf(".") -gt 0){
        $OpenFileDialog.filter = "$FileType ($FileType)| $FileType"

    }else{
        $OpenFileDialog.filter = "$FileType (*.$FileType)| *.$FileType"
    }
    $OpenFileDialog.Title = "Select the $FileType file that you want to use"
    $OpenFileDialog.ShowDialog() | Out-Null
    $File=  $OpenFileDialog.FileName
    $FileName = $OpenFileDialog.SafeFileName

    If($File){    
        If($FileName.ToUpper().EndsWith($FileType.ToUpper()) -gt 0){
            $WPFTextBox_Messages.Text = "$FileName selected. This is a valid $FileType"
            $WPFTextBox_Messages.Foreground = "Black"
            $File
        }
        else{
            $WPFTextBox_Messages.Text = "$FileName is not a valid $FileType! Select something else."
            $WPFTextBox_Messages.Foreground = "Red"
        }
    }
    else{
        $WPFTextBox_Messages.Text = "Nothing selected!"
        $WPFTextBox_Messages.Foreground = "Red"
    }


}

#endregion

#region Initialize

$RunMode = "Script"
$ThisScriptParentPath = $MyInvocation.MyCommand.Path -replace $myInvocation.MyCommand.Name,""
$ThisScriptName = $myInvocation.MyCommand.Name

#If the Script gets executed as EXE we need another way to get ThisScriptParentPath
If(-not($script:ThisScriptParentPath)){
    $Script:ThisScriptParentPath = [System.Diagnostics.Process]::GetCurrentProcess() | Select-Object -ExpandProperty Path | Split-Path
    $RunMode = "EXE"
}


#Set Version
$WPFLabel_Version.Content = "Version: $ScriptVersion"

#Clear some Stuff First
$MSIXData =$null

#Check Sideloading Staus

$CurrentSideloadingStaus = Check-SideloadingStaus

$WPFTexBlock_CurrentSideloading_Status.Text = $CurrentSideloadingStaus

If($CurrentSideloadingStaus.Contains("Enabled")){
    $WPFTexBlock_CurrentSideloading_Status.Foreground = "#FF0A7211" #Green
    $WPFButton_Change_SidelaodingStatus.Content = "Disable*"
}else{
    $WPFTexBlock_CurrentSideloading_Status.Foreground = "#FFFF0600" #Red
    $WPFButton_Change_SidelaodingStatus.Content = "Enable*"
}


#Check if the OS Version is supported
$OSVersion = (Get-WmiObject -class Win32_OperatingSystem ).Version

[int]$Major = $OSVersion.split(".")[0]
[int]$Build = $OSVersion.split(".")[2]

If($Major -ge 10){
    If($Build -ge 17763){
        $WPFTextBox_Messages.Text ="Your Windows Version supports all features of this tool."
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


#Check if user is Admin
If ([bool](([System.Security.Principal.WindowsIdentity]::GetCurrent()).groups -match "S-1-5-32-544")) {
    $Script:UserIsAdmin = $true
}
else{
    $Script:UserIsAdmin = $false
}


#Checkthe script is executed within ISE
If(($host.name).Contains("ISE")){
    $ISEMode=$true
}
else{
    $ISEMode=$false
}


#Do the Icon Stuff
If($RunMode -eq "Script"){

    # here's the base64 string of the image 
    $base64 = "iVBORw0KGgoAAAANSUhEUgAAAEAAAABACAYAAACqaXHeAAAdrElEQVR42u17B3gc5bX2O7OzvfemVe+S5d6NO9gGGQcIAdOSm/ATCCXAnwQCfxJCAjgFyL3khtw0UxyIgzHgChiMsS0JC1uWrGZJlrRalVUvu6vd1bb5z6wsYRNuLrnXjvP8f+Z5RrOzO7P7nfec8573zPeJwf/nG3OpB3Cpt38CcKkHcKm3fwLw9/iR2XPm2RiWUet0us6D7x0IX2qj/24ArFi1RtHd1flcTl7+TUqlSuL3jQV8vrHesdFRt0ql6giHw+0ajbojEU90pLhcHRKpdGDbSy9O/D8DQGpa+ivXfPGGzRtLr4RYLOH9fn/yBxmGQTQagT8wzoyOjGBoaAgjI0OJ4eHh8eGhob6xsbEu+rwnFot1qJQqt1gs7nA4HG4CyKvVavw//9nP+H94ACjsHQaT2f2zn/6EizES5kxbOwa7WhBqeh/ReJzn5Wawags4pQESmRIKhYI36A0wGPQMGQlOJEIkGmXGA+ME0CD6+vr4/v6+aK+3d7DH2+Ohgbv9gYBHxLJupVLZazAYPAajsVMmlY488cQT0UsOQGFR8XWr1q7b8fAjj/LjEwnG31mHD39yHbafGMFgkEeKwwxH/kzoWT/U8EOSCCdHE+JlfIhRMiKllVcaU2CwpmLdhlKcOdMKrU4Lg04PShuQ0RCJWCYej4NSCl5vL7q6OhMdHR1jbre7e2BgwM2yrMfn87nVanUbAVS+detW798NgLSMzKcfeviRB7LyisCIxMzgvsfxx8pBdA4EEGFlyL32IYw0VSAwEUcgEkdYYQMvUUAZHYARQ7CIRmGV+GCVjqOksBgqmQitDR9jLCJGICbl45yWosfCi5UmRqUxMeuuWIseAsFo0NNugIoAkkqlkEjEGBwcxOuvvx557rnnDuQXFDyyZ/fuUxcdALvTVfH8r3+9sKW9CyqVkmne/Qz8xjnCa3i4TIz0tEGjVkOhMUDvyoPCkg5fOAZ/MATfeBi+YBjxRIJ2Hs1uD7Jjp3CD7RQWpcbI8yKwHO105FlKFUYB87It6Hd/hKaqvZjgFWClRoiVVlgcWbj8io0Ehgxt7W48++yzoYMHD97a2FD/+kUDYPbc+apgKNz71FNPyfcfOMha5TEceecNeIJyMCyH+JybwAw0I8ZKEWPE5HkVlBryHIW4mgBSKhQU5hosnFkImVKFE00d2Lb1N4gPe3D3oiC+UDQBqYxLAiDsEn0WmJy7IR7Yg+jgscn3kwCxFH0iqJwrYS26HcFgEK++thM7/vynASLa9Iqyo8GLAkBObv5Cm9P10a233Jx4Y9ceNl/Wi/aBEPoTesRECiTW/xCJ8WEk/APgfV5EJLpkWsQi5PXoRPIYDY0jHh5HYUE+7EYtdv7iIaydmQanS481tjPITVVBTOEtGClxrMSEcR3k3t8hFvIizrOQy8XTIIgUTqQufgqjo2P46bP/hvbWFjTW1xfUVFedvigAUPm7u3TTtb+kWp84WXOKXSJvxJAvhEp/KsQyBcLLv43EmcMUygxFhAjirGVgTZnTxsdpjwT9ydcisRRjvR0Y2PUE7rrtBqgVQDQcons5LLacRIpFBjblC4ipZ8JX+yw6ogvg6RlDY/X7uH6tEbMLDZAZZ8JS/E1Un6rFlp//Ah53W3w84E+pra7qvSgAOFJc2x559Hs3fXD4KCjUmGu0x6HjxvGH9nRo1nwHnDkTH1Q3g4sHKVdV4ORqSJUayBQqKokKiClfBQOFXCfRAP9gDwbKX4NWweI2yyHMzyZyk056X6WUIGy/je5TofnkXkQ1S0AVAI2NjRgb9uLOL0hQsuQmaFM3Yedbe/DSH19FbfUJj0ajTT918jh/wQGYOXse09ff37Jly5bMba9uZ8RMDLeZKiDmgB3SryBl5dfRHJThnV3bUarvRuWh3ZBnzEaQ02MEBkQpHTjSBVKFEkZnBkVFBNGJMNIlA7ic34XZ5mG4bKrp/Gc4KQZNd8Ao7sbhigZAkYGqqiokiEAnAr34+nVqzFr5bcgNM/GLXz6Pw0fL8dHRQ7sH+/uuFsZ7wQEoLC5xJsB6Hnzgfmb7jp2MXTKGLxhrEUiI8ZLrV1gwbwHeONGGtsr3MJNzw9W3G9cvNcKikyKSYDAyIcJARA5vWIHemAH+mAS54k7MTYnAYSbBJOR20vDJ/E5IbfAqrocNlfjzwWFKGSWOHTtG1UYFLYH20P/KhnXmYwSqEQ8+9Cg8nk58XHHksaGB/h9eFAAcTtd1M+fO37FwwYLE4aNl7Fx1FxbpOnE8koeWuT+D0paOXfv2EQmOIqd/D0qtLVhaoEWCBh6TWaGM90yXuRYsgyt7FjTdvzuP1dlzymBQWox+bimUY7vx1jEpent70dbWRiQox2UlCdxxyxzIMh7BRCSGO+99AIL0bm6s3dTn7dl1UQCw2uxP3XjzbQ8NDA7x3t4+ttRyGpnKMWzjb4Rpye2o80lQfeRtQSdgXvUDKC2REMvLUMsvAafPhitaARPbQ4ZyOMlfjaI0BdSDb0waftboZAQQGMJ5P7sEIUkBop5t2FulQU1NjSCbYTUb8N07UoWSDM72VXR4PPjBj7fA2+WJjw4PZ3rcrZ6LA4Dd+d7d935z1YdHjjKJRJz5qquamB74lf5nyCtZiEM1LUA8Bie6saDpB1g/zwgRx6F8YhWJGglma5tg4gYQ5JWoF23CHKsXikDltNEiTvRJFNB5c2Q1VGo9/J49eKNCiuMfH6PGahRzi8144ltzoLYtBdRrUVVdg1//biua6mv7ObHYfqapIXHBAcjNL5SGJyK9d951l/adAwcZjTiCr6bXYl9gNrpnPATOlI5DB/Yhs2g2sluexzLREZRkqBFiNOhTLkWYeKKYOUjylUVn0Ixh3VWYpToGadSDUFyMupEMzLJ6oFJMAkAJj+OjVyDTEgLjq4Bt1mbs2vEgKo6GsHBuHm7clIdxyVriiSLse/sA9r79Dqoqyw8TAa6YGvMFBSA9M3u2wWQ9sWHDehwtr2Dy1CNYbOrBD6MPYv78pTgTEKH2vddgIGBuZLdjbYkGGpUYA9HZSLxZA+mNa2BTVCSNOzmUAmnqepSwOyFiJtAbcaK2z4hlKU3Qqblk+Id4DaqHl2Cmk3gjTn1A3gK8+Id7EAlk4frSOTCZlOjnr4Vc5cIvnnsejU1NqDpW/ovhoYEHLgoAFpv9G5etXPvv1H3xpLuZNbZuDPAqvG+9HzkFM3GChEh3QyVyJqpxZ04L5uZqkuHsfzeGgChGkjULpnmj4MQivNeZi/S8+Sjgd5KxLLqYpRiLWZEl2gutahIAb9CBM/4ZWJJWh65hILNkFUXYdxEPmlC6tpgqhQwdoetg0JvwwHcewcjwMBrrqr9MFeCliwKA0WT5IxHg5pbWVoRCYebL2a2wKCZQ6UvBGS0Nri2KviAHSXwcP07bgwV5WoQCCcSOxMAWipAYlsK6WpHM9bdaSzCvyIFM9sPk+aDmS4glJEiJ7aDInyTE+v4MjCRysDKjAoerx3CgSo4UuxZ5jm5sWOGklDKiw78GHAH46A9+BDKc7+/1lni7PXUXBQCbM/X09V+6Maf61ClGwiaY+4qbIRazSbISBiyw4UCIQ/+EDCZJEKlWOeK1IQSMOoiiJH9ZFVJncvBTy/tO11ysLuZhF9VRdydFv+oGxMc7kK8+Pk2CB1uyodVZsSj9BPwhoG54LbTRNzGjwJxUiv0Bolr/7OQTp1//9g9wt7b4SCBZOtpaph+7XTAA0rNyTGKxtHv12su5+oYG1qkM48sFnumydW7p+uRchOH9QSjXaeCvozGlaJCaxeJ0rxQt0cW4PLeLJLQXQ3wmesakkCvUKDLW070ceEaEnVU5KMmRothxBk2dLIKSxUhXHYbNpkmC3tKbjeFQBuV+M97avRd11cc/IgJcfO64LxgAVrtjQ1Zu4b7MzEy+3d3BLLCOYl360Nn6zZ4PxNnzoD+BgSoGuRsUcB8KQTFLDaedw3uNUsSNK+j+4/DV96DvPQn010mBVilc66XUN0jgC0uxpzodq2bHkGr0orYrHcFwDAsL+qFQSJLff6y5gBRgGvbsewfHq06i7uTHvyEC/PpFAUBvMP1o5dp1j4bDExjz+Zgv5vahyBz6i9586lxE5x0VEXDpaooeyvH9Q5CuNMNolGBbuQIpuYuwWHEIilc96F+YBvsSFYmJBoxelQdHiRpNXgWOu13YuMiHBE8V0XIL/rjtGdxzoxoKAohQxrvHi+BKzcJPfv4sqD9BY+3J+0aHh567KACYLPb3r9hw1apeUmGRiQnmm3MpfBX8J4Zzn4iXqfCv2RFB8fUK8hiHoV0j0G6yQCrh8K/vabF8QQ4cb7wGiSsDTmpoBOILvFSLcE4e7Gs0OFSrQA+11zet7EH7kA0nm+UI+07iX64zQiaTwE8RcuDjTGSmp+PRx36MsdFhQQUuG+jzll1wAFJSM6RklHfN5et0kUiEYSMjuDWr7rx8Zz+V/73tEXTXibD4BmUSjLY/DSDzFjsCERH+UGbFEicH5/46WO4yw5CiTN7jO9CG4IACrjvSse2AFCqdC9cv70ZffD3eL2tFlrmFgFMkr233KlFxygqT0YBn/u2XgvHkl7C12+MeveAAmK32AovdWb9s6TJGr9dBH6Ear6g+L/cFI0XnnNfs8VGTYkDxfA7BEI8PnhvFxu/Z0NjN4IgnE0vH22AYCyPnDut01Iwd6oCPCljqt/Lw7J/FWDjbhmVF/fDJvoZ3Dx7GqpIW8rh6Mv9rlWh068CyLLa9sh0tjbWN/b09hZ8e+wUBQKPTL83MKTianp6BtatXQDdWDkf0xDk5P2mA6BwSLHt5DM4VBuTmSdDjCaNh/wQ2PGjF/hNx9ESLUOovh5pRwvhF8zRofVtPIq5Nh2SDFb/dLca1lxuojQ5CarsdL778H7j/Fg5KpTQZaW++L0cwYk0+Li8/VonjH5W9QgDcfFEAKCgsNoYmou0OV5p61swZmJ2uQOr4XsjEOK91nfKkcB7nmckHmlIODWVEZBIZ5qzQ4rfvxiE3FmBjogxsnxyWW+1JIIP+KDyP1UF3az4GjGq887EKX7tGjihjgj8xhwhuN265Wk4qkkuW219t45CRUYiPKj9Gd3cPDry991v9fd6nLwoAy1estFdWHnty9ryFX9EbzfD5fBDTIErXLoKCHQc70Y94qB862QSsOgZaJUsDZc/hBi55pD94fHsC80sycOWMHoRGEzC4lEnAOj7oxsixBGb9n0zsKQuhuceC/30rg4HRKCprY5hVpEJx3uTDkkCIxb9ujeOypUtw6HAZPB43Pio7csWZluYDFxSAq64qXej3+x/bvHnz2oLCQk54T5ixSfA8urp64OnsRGB8HGJOjBi1wIcOHxUkMuRUpSxaBnYDC4dBBIdJDIt+UjE+t4fFtaspt/PHziNOIWKikQTUejl+/uIg5OpU3H8zP11i2XPa5GZ3HC9ThSm9ch2OnziJY8cq4h53e1pFeVn3BQFgxYoV8xUKxfe/fuedG0xmC3uqoYmpOnECDfV16OxwQ6vVISsrCxmZmbDb7bzZbKL6bqQSJ0FvXz+6e7wYHR1lBHACgXGMjI5ibIzSIBEn6SxGul2KbAcLp5W0v01KRyl0WgnEEuoXwOKbT3Vj1dIsbL6KmSbYc8Ha9e44yo/HcPONX0J942m8/OIL3XKFPPWtN3Ym/kcAXHHFFXOJVR+55557NnHUYfx+6wvMRCyB/MIi9PX2oqXpNDra25JfKxJLKMzFkzsnHCXQ6fRISXHCbrPyQqQI01ZiCn+xMBEaiVA1CBIwPmZiYoIiJk7REsLg0DCMBh1mzciDRsmDo5T6oLwbc2cosX65AjaLjISPeNr4prYJPP/CEByOFFxdeiV6COxHv/vwvuPHP77qs2z6XACsWbOmmA4/vPfeezfpdDrR448/jkOHDgnzf5g5ey4WzC8h9ReCl34sFAzh1KlTCJNBIu4sAAIYZ0GYAkRrtGDWsnWwONLw0cFdkI738yoSOxzDI046f4LUXZyGJ1zv9wcwHgwyQgTpdTphBhktrW3JiQ6h2bJblEhNUVHqsWj3RJNzgl//2r8kp+GE6Lr7rru+X1NT/aO/GYBVq1cXx6LRH9x///2blCoV98zTz+DAgXcZnnJcmOMvLV2DJ598FuLxJmLZP8FS8AVqXv6Eh3/sQY93AGryOMsJxkumgZBSQ1OyYiPWbLwRoWgMv3/lNZyqrcVSeQ+4sA+iCJU1iRRihYo0vwoylQYShZIXQAlHokRwoeQUF3ieUSoVvJAyIpZlhMdqSoUc6WmpWL1yORx2O4RI+t73vjdMpFy8ffufvJ8bgFWrVhVFo9FH7rvvvuu1Op14y5YtPHmc4ROJZC3PIM9vujpB5KXF+k0vgolP4HTFy+ge6EaTZxQdXjnMNgdc4Q9QnG1E66gUxwetGFaUoHjxOmLrLJyoqcGrr78FTqlHWn4JnPXbIAqPTQMgESZJFMrkUSYc5crpc2HOgJPJkxOn46EQHwsHMdDTmYw8m9XCrF6+hBexTGL37t01fX19dxsMhkphoqSzs3NqYQX/mQCsW78+LxAIPHLP3fds1un13NNPP40PDr6fnIOf2nJyconYbPjGNRwO76uFIvcqLF59FV559TXK7xTk5+ch1a7D7g/L0TXBoWbMiDFei3SnHUXpTsiYGA5+eCjJGba8uSguSkXx0JvwU60e6RyGbqgXAaUOlCPJCPjPAJg8KpJHq8kIg0KUeOLxH/HkdbalseZAOBS+nSJkgMYurDzhhdLMC6E7abywf/JQ9Kabbsro6ur64R133HGjze7gtm7dyry9fz/iYerm6NrR8WDyahkVumvnkMeMCsgpfNNMUoy5LoM4UEN56gMnkUOvZhGSW/F04GvIcTmRajVCr5IT2TF4Y/97aDldD7lKC9ecVSg0+PAV3XbkOIgMpWK0vDeI8SMDYIZ64Av60SPVYaCgGJKMDLASaoiknw3ASL839vZbryMRj4luu3kz89MtT7xZV1t78znePtfzginTK0iY9es3mLOzs04/88wzhsGhocR9996baGhoEA319TFZWZkYovKkoBJjiZzBNUUsNT5FKOAzoRjsgT4RxjAvRtysRxc/jBHtKBgiqvvbr8WgOB2FaZNez3Ea8efXX8epplboXTmwFSxC0cRBfMn0AXLSNAScGKff7QAXjYInYBOhBML94+C9o1B/5ym8+ebrYAmgzbdtRoSXo6d3FAmWeEXJ8IeP1EVrT1QSxYjZVFcKrt1USqz/0Ov19fVf/RQAwi6E8nmr1JjVq9ekZ2Rlt5WWljIWiwVajTo2PDQU7+lwMxqjSfL22/txsroa6wvtfFbAy6wxGFHZNYaeMS+WsGJs7RzA9XPykE61vz7ahD5i5O+2XwMLdWEpdivyc3Jw4HA5vMRb5uxZ4BMxzO3/D1yV1oX8LH2y7o+Smit7sh3iiRAMCR8sIgKdD6Jm+ZVoIN4RizTU5KRRYyMiLSDB0PBIMiVOtjcnPN0hhkozI5NJYTaZcA0B8PB3vvXa6cbGOz/lecHr439BgqtWrXHYbfaOzOwcbozKjTBdatDrhVqNFKcDJvpStUrJi6N+XuLthvR0E1stEycmWhrwUflJNqFK4AarCZaWVtTkJ1ANBf59uJQkrwpmOTCjsAD1kiJ4RwJQB9woGXoJa7MDKMg0IMjLsL97DsRkLNPagWJ9GBpxjMgwBjlikK26HN1BDi0tA4hHE6QbJDCTk8qq6lDf2inM//H+QCBZlUiXwGIWANiIbz94//bm5ub7zjFe8Lr/M6vAkiXLTEReXRmZOVJPVze6iZxisViyjRRQNZGCs1jMsJrNSfIThIzQY2voxxn6bo1SgsjJE8zA8YOJLcdOsPuO1CUHk5IpPNaegez8GQiySsywSTCnMBt6ZQLRxp+CYRn4Ygoc8S9Bl1+HzjEFRgMTuGO+n/ilAU4LQyKHo5pOHqGxjIzxGBpJ8H/4cxf6fTxjpTG1uzswZfwkAGZKgY24/767X2lra/v2WRv9/5nxSQAWLVyscTidPdm5+cpO0u8CAALrC3V++qLp13zSCynOlJhGoxIJJcfpcCA91TUpdQkw0g3UB3TB7W5Hf/8A/D7hkVUiKWBcLhcQ86OI24u9vZk4NpwJJctRw5RAPMZAzUWwokiPfQdOkkT2ITs7DzabAL6GgOvkX915jNcZs9kQ6YCG001JJ00ZL+wCKJs2Xok7vvbV3/f39//orPeHPiv0p21bsGCR3Ga1efMLi7We7m4imL6/AsDkduttX4nU19VwVVUnWeEz4fmeTquFk1LGRRHionIoEJLwcIRcRIIkIiyUSK71Cw3WxtP0QaayT8U0juuJksUMw5qZaEyJFal+aORyijyZIG6Sc4rjAT9qqL53dHqwbMkilFVU4Exr+3mGn915u80W0yilomef/rlAgEdpH6R9DH9lY+bNW8AR+fXOKJltPBcAAd1zAZjcJyvn5Vesi55pauTcnk6BgCbfp9FOrQIV3tCo1TxxEwkTK6kyKyPwSVpqajJ9hNVfvjEfPzg0yPuGu5LreuQiP29Qhcl4lhUrHBArMxCMG/nyjxsY0iZMQX4utr70MoaoNxAMFlaQiMWTR+E8HovGJZwo/s6+3c8kEokXxGJJdzQaCeC/2JiZM2exKU5Xz7wFi61dXi96BwaSqyvYSaswaQ+bVAzJ16QMcnNzou2tLVyMp0+YSU8J709eO3nfZcuXo+0MsXRnV0IilXJCJJAmgQCKyWQg7Z5CgKQgIy0NVqsFpDzhJfC7unt4Ei7JZkir0TB5udmk5314/Iknk83RlMfPA4DUaV1NVV9Pl0cgvlphaS19X/C/Mn7SnbSRtN21dNmKjQxpdj+JHqHUCD8y6f1PjJqKBKlYFOtobxXZXRnTEcCc/Vw4Ea5Oo7J1puk0Tx1RnEKaY84DdPJaQZfFYtFkG6wm1edw2Ij4zMmosZGWN5qtyabntR07UH70cDIyzw174qMJPsFLGutruvt6ur8jGE9SvYWu+9wLrpMA2Gz29N5e75fp5tkKpSpLo9WnmixWldXmYPVU96kRSubl1MB7uztjno520YIlyxmhNk+DdDZNBAhEdGw+3cC7MrLicrmCmwJGYP9PwJwcAjt5U/K+8EQYfRSJ8xcuxrLLVgjLX/Hu23tRXnbkPABErIh3OmwU8nuafWOjD9P3tVI8NhGY8c9r/DQAU5tcoVTS/VZibQuxeSGVmAUSqSxXpdZkmswWu9OVKnGkpMJ9pinhcbcxt97+jbMRkPT/2TSZfN3d6cbphjp+2cq1cYViCoCzRn/6+rMkO/W6r68XwgPWpQRAJ5Hfu/sFAI4mI2UKgIDPxzfWVe+g8f6e7nHT3kyp+zevIv+r7TB5TpTgeSP9sJ2+PD0ei80igwvoI8p+hp+/+LIbUtMzWAKGyqCZrpdNe/dY2Yd8fW0NydfbEyq1SjT1c0IEMFOvP2X41OtOTwfsDuckAB433tm/DxXlRwRuipHxouB4AC2NdTvJSb+h61vIUe1/q+GfC4DP2jhS3VKpNEZRkkGd1zLywAIinQK1Vu9yOFNSMnPyZJmZ2cyJyopE3amq2GNPPiNRazQ4LwKA8/jiXH4Q/ghPlgTFR7yEjg53MgXKjh5O2MymeEtLC+tubX6BwNhJzqiiY+9/1/j/FgB/ESUKJY2foYIPMxGaKzIxMYtycY7AcfR+G/HEo0UzStgsEjUkuEDpNM0TmD5+wgcCGDXVVaQhDBQBy5OCSgDg6IcfQC6RhCo/KvslXXiIjC8n40f/Z6O/SIulqYqIpFJZgkjVEBwfX0n1eB5FTolGp8tMcaW5cvIKFCS8mKzkswVnUnILyTvFH0ePHEoCsGTZZXC3t+HtfbuxfdsL3aPDQ8LaPg/V+CP0nZ+rzF0SAD5r02i1StIB1lg04qIaPZO8N18mV6QbzeaM3LxCa8msOZzwcFVQjSerjmPBosWUAstRdbwSP/7+dztIU/yKAKoglCp44oILNa5L9m9zKpWaNAyno9Jmj8djmaTvFxDBzVJptJkEls1kMgt9hbexoa6KyHefTCYvCwbHPRd6HP9Q/zdIGkROFcdGpKonfcESu0co13tHR4b7L9Zv/kMBcCm2fwJwqQdwqbf/C3pgOfVSQHzMAAAAAElFTkSuQmCC"
 
    # Create a streaming image by streaming the base64 string to a bitmap streamsource
    $bitmap = New-Object System.Windows.Media.Imaging.BitmapImage
    $bitmap.BeginInit()
    $bitmap.StreamSource = [System.IO.MemoryStream][System.Convert]::FromBase64String($base64)
    $bitmap.EndInit()
    $bitmap.Freeze()
  
    # This is the icon in the upper left hand corner of the app
    $Form.Icon = $bitmap
 
    # This is the toolbar icon and description
    $Form.TaskbarItemInfo.Overlay = $bitmap
    $Form.TaskbarItemInfo.Description = $Form.Title
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
    Use-Filter
    $WPFTextBox_Messages.Text =  "Got the list of Appx/MSIX installed software"
    $WPFTextBox_Messages.Foreground = "Black"

})

$WPFButton_Filter.Add_Click({
    Use-Filter
})

$WPFButton_Open_Manifest.Add_Click({
    Open-Manifest
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

$WPFButton_Start.Add_Click({
    Start-App
})

$WPFButton_OpenRegedit.Add_Click({
    Open-Tool -tool "regedit.exe"
})

$WPFButton_OpenNotepad.Add_Click({
    Open-Tool -tool "notepad.exe"
})

$WPFButton_OpenCMD.Add_Click({
    Open-Tool -tool "cmd.exe"
})

#Edit Manifest Actions

$WPFButton_EditManifest_SingelFile_Browse.Add_Click({

    $SelectedFile = Select-File -FileType "msix"
    If($SelectedFile){
        $WPFTextBox_EditManifest_SingelFile_SelectedMSIX.Text = $SelectedFile
    }
})

$WPFButton_EditManifest_SelectFolder.Add_Click({
    $SelectedFolder = $null
    $SelectedFolder = Select-Folder

    If($SelectedFolder){
        $WPFTextBox_Messages.Text ="The Folder $SelectedFolder was selected"
        $WPFTextBox_EditManifest_SingelFile_SelectedMSIX.Text = $SelectedFolder
        $WPFTextBox_Messages.Foreground = "Black"

    }
    else{
        $WPFTextBox_Messages.Text ="No Folder was selected"
        $WPFTextBox_Messages.Foreground = "DarkOrange"

    }
})


$WPFButton_EditManifest_SingelFile_Edit.Add_Click({

    $SelectedPath = $WPFTextBox_EditManifest_SingelFile_SelectedMSIX.Text
    If($SelectedPath){
        Edit-Manifest -Path $SelectedPath
    }
    else{
        $WPFTextBox_Messages.Text = "Nothing selected!"
        $WPFTextBox_Messages.Foreground = "Red"
    }
})

$WPFButton_EditManifest_SelectCertToImport.Add_Click({

    $SelectedFile = Select-File -FileType "pfx"
    If($SelectedFile){
        $WPFTextBox_EditManifest_SelectedCert.Text = $SelectedFile
    }
})




#Change Signature Actions


$WPFButton_ChangeSignature_SingelFile_Browse.Add_Click({

    $SelectedFile = Select-File -FileType "msix"
    If($SelectedFile){
        $WPFTextBox_ChangeSignature_SelectedMSIX.Text = $SelectedFile
    }
})

$WPFButton_ChangeSignature_SelectFolder.Add_Click({
    $SelectedFolder = $null
    $SelectedFolder = Select-Folder

    If($SelectedFolder){
        $WPFTextBox_Messages.Text ="The Folder $SelectedFolder was selected"
        $WPFTextBox_ChangeSignature_SelectedMSIX.Text = $SelectedFolder
        $WPFTextBox_Messages.Foreground = "Black"

    }
    else{
        $WPFTextBox_Messages.Text ="No Folder was selected"
        $WPFTextBox_Messages.Foreground = "DarkOrange"

    }
})


$WPFButton_ChangeSignature.Add_Click({

    $SelectedPath = $WPFTextBox_ChangeSignature_SelectedMSIX.Text
    If($SelectedPath){
        Change-signature -Path $SelectedPath
    }
    else{
        $WPFTextBox_Messages.Text = "Nothing selected!"
        $WPFTextBox_Messages.Foreground = "Red"
    }
})

$WPFButton_ChangeSignature_SelectCertToImport.Add_Click({

    $SelectedFile = Select-File -FileType "pfx"
    If($SelectedFile){
        $WPFTextBox_ChangeSignature_SelectedCert.Text = $SelectedFile
    }
})



#Cert Actions

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


$WPFButton_SelectCertToImportFromMSIX.Add_Click({

    $WPFTextBox_SelectedCertFromMSIX.Text = $null


    $SelectedMSIX = Select-MSIX

    $WPFTextBox_SelectedCertFromMSIX.Text = $SelectedMSIX


})

$WPFButton_InstallCertFromMSIX.Add_Click({
    Import-CertToTrustedPeopleFromMSIX
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

$WPFButton_StopServices.Add_Click({
    Stop-Services
})

$WPFButton_Change_SidelaodingStatus.add_click({

    $CurrentSideloadingStaus = Check-SideloadingStaus

    If($CurrentSideloadingStaus.Contains("Enabled")){
        Disable-Sideloading
    }
    else{
        Enable-Sideloading
    }

    $CurrentSideloadingStaus = Check-SideloadingStaus

    $WPFTexBlock_CurrentSideloading_Status.Text = $CurrentSideloadingStaus

    If($CurrentSideloadingStaus.Contains("Enabled")){
        $WPFTexBlock_CurrentSideloading_Status.Foreground = "#FF0A7211" #Green
        $WPFButton_Change_SidelaodingStatus.Content = "Disable*"
    }else{
        $WPFTexBlock_CurrentSideloading_Status.Foreground = "#FFFF0600" #Red
        $WPFButton_Change_SidelaodingStatus.Content = "Enable*"
    }


})


$WPFButton_EnableDevMode.add_click({

    Enable-DeveloperMode

    $CurrentSideloadingStaus = Check-SideloadingStaus

    $WPFTexBlock_CurrentSideloading_Status.Text = $CurrentSideloadingStaus

    If($CurrentSideloadingStaus.Contains("Enabled")){
        $WPFTexBlock_CurrentSideloading_Status.Foreground = "#FF0A7211" #Green
        $WPFButton_Change_SidelaodingStatus.Content = "Disable*"
    }else{
        $WPFTexBlock_CurrentSideloading_Status.Foreground = "#FFFF0600" #Red
        $WPFButton_Change_SidelaodingStatus.Content = "Enable*"
    }


})

#AppAttach Actions


$WPFButton_AppAttachConvertToVHD_SingelFile_Browse.Add_Click({

    $SelectedFile = Select-File -FileType "msix"
    If($SelectedFile){
        $WPFTextBox_AppAttachConvertToVHD_SingelFile_SelectedMSIX.Text = $SelectedFile
    }
})

$WPFButton_AppAttachConvertToVHD_SelectFolder.Add_Click({
    $SelectedFolder = $null
    $SelectedFolder = Select-Folder

    If($SelectedFolder){
        $WPFTextBox_Messages.Text ="The Folder $SelectedFolder was selected"
        $WPFTextBox_AppAttachConvertToVHD_SingelFile_SelectedMSIX.Text = $SelectedFolder
        $WPFTextBox_Messages.Foreground = "Black"

    }
    else{
        $WPFTextBox_Messages.Text ="No Folder was selected"
        $WPFTextBox_Messages.Foreground = "DarkOrange"

    }
})


$WPFButton_AppAttachConvertToVHD_Convert.Add_Click({

    $SelectedPath = $WPFTextBox_AppAttachConvertToVHD_SingelFile_SelectedMSIX.Text
    If($SelectedPath){
        ConvertTo-VHD -Path $SelectedPath
    }
    else{
        $WPFTextBox_Messages.Text = "Nothing selected!"
        $WPFTextBox_Messages.Foreground = "Red"
    }
})


#endregion


#===========================================================================
# Shows the form
#===========================================================================

If($ISEMode){

    $Form.ShowDialog() | out-null
}
else{

    $Form.Add_Closing({[System.Windows.Forms.Application]::Exit(); Stop-Process $pid})

    # Make PowerShell Disappear
    $windowcode = '[DllImport("user32.dll")] public static extern bool ShowWindowAsync(IntPtr hWnd, int nCmdShow);' 
    $asyncwindow = Add-Type -MemberDefinition $windowcode -name Win32ShowWindowAsync -namespace Win32Functions -PassThru 
    $null = $asyncwindow::ShowWindowAsync((Get-Process -PID $pid).MainWindowHandle, 0) 


    # Allow input to window for TextBoxes, etc
    [System.Windows.Forms.Integration.ElementHost]::EnableModelessKeyboardInterop($Form)

    $Form.Show() | out-null
 
    # This makes it pop up
    $Form.Activate() | out-null
 
    # Create an application context for it to all run within. 
    # This helps with responsiveness and threading.
    $appContext = New-Object System.Windows.Forms.ApplicationContext 
    [void][System.Windows.Forms.Application]::Run($appContext)
}