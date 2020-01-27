# ezfio 1.0
# earle.philhower.iii@hgst.com
#
# ------------------------------------------------------------------------
# ezfio is free software: you can redistribute it and/or modify
# it under the terms of the GNU General Public License as published by
# the Free Software Foundation, either version 2 of the License, or
# (at your option) any later version.
#
# ezfio is distributed in the hope that it will be useful,
# but WITHOUT ANY WARRANTY; without even the implied warranty of
# MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
# GNU General Public License for more details.
#
# You should have received a copy of the GNU General Public License
# along with ezfio.  If not, see <http://www.gnu.org/licenses/>.
# ------------------------------------------------------------------------
#
# Usage:   ezfio.ps1 -drive {physicaldrive number}
# Example: ezfio.ps1 -drive 3
#
# When no parameters are specified, the script will provide usage info
# as well as a list of attached PhysicalDrives
#
# This script requires Administrator privileges so must be run from
# a PowerShell session started with "Run as Administrator."
#
# If Windows errors with, "...cannot be loaded because running scripts is
# disabled on this system...." you need to run the following line to enable
# execution of local PowerShell scripts:
#       Set-ExecutionPolicy -scope CurrentUser RemoteSigned
#
# Please be sure to have FIO installed, or you will be prompted to install
# and re-run the script.


param (
    [string]$drive = "none",
    [string]$outDir = "none",
    [int]$util = 100,
    [switch]$help,
    [switch]$yes,
    [switch]$nullio,
    [switch]$fastprecond,
    [switch]$quickie
)


Add-Type -Assembly System.IO.Compression
Add-Type -Assembly System.IO.Compression.FileSystem
Add-Type -AssemblyName PresentationFramework, System.Windows.Forms
Add-Type -AssemblyName PresentationCore

Chdir (Split-Path $script:MyInvocation.MyCommand.Path)

function WindowFromXAML( $xaml, $prefix )
{
    # Create a WPF window from XAML from DevStudio
    $xaml = $xaml -replace 'mc:Ignorable="d"', ''
    $xaml = $xaml -replace "x:N", 'N'
    $xaml = $xaml -replace '^<Win.*', '<Window'
    $xml = [xml]$xaml
    $reader = (New-Object System.Xml.XmlNodeReader $xml)
    try{ $window = [Windows.Markup.XamlReader]::Load( $reader ) }
    catch { Write-Host "Unable to load XAML for window."; return $null; }
    # Set the variables locally in the calling function.  Please forgive me.
    $xml.SelectNodes("//*[@Name]") |
      %{Set-Variable -Name "$($prefix)_$($_.Name)" -Value $window.FindName($_.Name) -Scope 1}
    return $window
}

function CheckAdmin()
{
    # Check that we have root privileges for disk access, abort if not.
    if ( -not ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator") ) {
        [System.Windows.Forms.MessageBox]::Show( "Administrator privileges are required for low-level disk access.`nPlease restart this script as Administrator to continue.", "Fatal Error", 0, 48 ) | Out-Null
        exit
    }
}

function FindFIO()
{
    # Try the path to the FIO executable, return path or exit.
    if ( -not (Get-Command fio.exe) ) {
        $ret = [System.Windows.Forms.MessageBox]::Show( "FIO is required to run IO tests. Would you like to install?", "FIO Not Detected", 4, 32 )
        if ($ret -eq "yes" ) {
            Start-Process "https://www.bluestop.org/fio/"
        }
        exit
    } else {
        $global:fio = (Get-Command fio.exe).Path
    }
}

function CheckFIOVersion()
{
    # Check we have a version of FIO we can use.
    try {
        $global:fioVerString = ( . $global:fio "--version" )
        $fiov = ( . $global:fio "--version" ).Split('-')[1].Split('.')[0]
        if ([int]$fiov -lt 2) {
            $err = "ERROR! FIO version " + (. $global:fio "--version") + " is unsupported. Version 2.0 or later is required"
            if ($global:testmode -eq "gui") {
                [System.Windows.Forms.MessageBox]::Show( $err, "Fatal Error", 0, 48 ) | Out-Null
            } else {
                Write-Error $err
            }
            exit 1
        }
    } catch {
        $err = "ERROR! Unable to determine FIO version.  Version 2.0 or later is required."
        if ($global:testmode -eq "gui") {
            [System.Windows.Forms.MessageBox]::Show( $err, "Fatal Error", 0, 48 ) | Out-Null
        } else {
            Write-Error $err
        }
        exit 1
    }

    try {
        $out = (. $global:fio "--parse-only" "--output-format=json+")
        if ($LastExitCode -eq 0 ) {
            $global:fioOutputFormat = "json+"
        }
    } catch {
        # Nothing, we can't make exceedance
    }
}

function ParseArgs()
{
    # Set the global values to the param() values, so that Parse() can see them

    function IntroDialog()
    {
        # Gets user test parameters if not specified on the command line
        $xaml = @'
<Window x:Class="Window2"
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="ezFIO Drive Selection" ResizeMode="NoResize" Height="281" Width="456">
    <Grid>
        <Label Content="Drive to test:" HorizontalAlignment="Left" Margin="24,16,0,0" VerticalAlignment="Top"/>
        <ComboBox x:Name="driveList" HorizontalAlignment="Left" Margin="115,20,0,0" VerticalAlignment="Top" Width="218"/>
        <Button x:Name="startTest" Content="Start Test" HorizontalAlignment="Left" Margin="358,79,0,0" VerticalAlignment="Top" Width="75"/>
        <Button x:Name="exit" Content="Exit" HorizontalAlignment="Left" Margin="358,125,0,0" VerticalAlignment="Top" Width="75"/>
        <GroupBox Header="Information" HorizontalAlignment="Left" Margin="24,55,0,0" VerticalAlignment="Top" Height="109" Width="309">
            <Grid>
                <Label Content="Model:" HorizontalAlignment="Left" VerticalAlignment="Top" Margin="2,0,0,0"/>
                <Label x:Name="modelName" Content="Happy NVME" HorizontalAlignment="Left" Margin="50,0,0,0" VerticalAlignment="Top"/>
                <Label Content="Serial:" HorizontalAlignment="Left" Margin="8,25,0,0" VerticalAlignment="Top"/>
                <Label x:Name="serial" Content="NVME001" HorizontalAlignment="Left" Margin="50,25,0,0" VerticalAlignment="Top"/>
                <Label Content="Size:" HorizontalAlignment="Left" Margin="15,50,0,0" VerticalAlignment="Top"/>
                <Label x:Name="sizeGB" Content="100GB" HorizontalAlignment="Left" Margin="50,50,0,0" VerticalAlignment="Top"/>
            </Grid>
        </GroupBox>
        <GroupBox Header="WARNING! WARNING! WARNING!" HorizontalAlignment="Left" Margin="24,176,0,0" VerticalAlignment="Top" Width="398">
            <Label >
                <TextBlock TextWrapping="WrapWithOverflow" Width="376" HorizontalAlignment="Center" VerticalAlignment="Center">All data on the selected drive will be destroyed by the test.  Please make sure there are no mounted filesystems or data on the drive.</TextBlock>
            </Label>
        </GroupBox>
    </Grid>
</Window>
'@

        $intro = WindowFromXAML $xaml 'intro'
        $intro.Icon = $global:iconBitmap

        $pd = @{}

        $intro.add_Loaded( {
            $intro.Activate()
            $intro_driveList.Focus()

            $global:physDrive = $null
            # Populate the physicaldrive list
            $drives = Get-WmiObject -query "SELECT * from Win32_DiskDrive" | Sort-Object
            foreach ( $drive in $drives ) {
                $idx = $intro_driveList.Items.Add( $drive.DeviceID )
                $pd.Add( $idx, $drive )
            }
            $intro_driveList.SelectedIndex = 0
            $drive = $pd.Get_Item( 0 )
            $intro_modelName.Content = $drive.Model.Trim()
            if ($drive.SerialNumber -ne $null) { $intro_serial.Content = $drive.SerialNumber.Trim() }
            else { $intro_serial.Content = "UNKNOWN" }
            $intro_sizeGB.Content = [string]::Format( "{0} GB", [int]($drive.Size/1000000000) )

            $intro_driveList.add_SelectionChanged( {
                $idx = $intro_driveList.SelectedIndex
                $drive = $pd.Get_Item( $idx )
                $intro_modelName.Content = $drive.Model.Trim()
                if ($drive.SerialNumber -ne $null) { $intro_serial.Content = $drive.SerialNumber.Trim() }
                else { $intro_serial.Content = "UNKNOWN" }
                $intro_sizeGB.Content = [string]::Format( "{0} GB", [int]($drive.Size/1000000000) )
            } )

            $intro_startTest.add_Click( {
                $idx = $intro_driveList.SelectedIndex
                $drive = $pd.Get_Item( $idx )
                $global:physDrive = $drive.DeviceID
                $intro.dialogResult = $true
                $intro.Close()
            } )

            $intro_exit.add_Click( { $intro.Close() } )

        } )

        $intro.ShowDialog()
    }

    # Parse command line options into globals.
    function usage()
    {
        # How to use the script, and some handy info on current drives
        $scriptname = split-path $global:scriptName -Leaf
        "ezfio, an in-depth IO tester for NVME devices"
        "WARNING: All data on any tested device will be destroyed!`n"
        "Usage: "
        [string]::Format("    .\{0} -drive <PhysicalDiskNumber> [-util <1..100>] [-outDir <path>] [-nullIO]", $scriptname)
        [string]::Format("EX: .\{0} -drive 2 -util 100`n", $scriptname)
        "PhysDrive is the ID number of the \\PhysicalDrive to test"
        "Usage is the percent of total size to test (100%=default)`n"

        "`nPhysical disks:"
        $drives=Get-WmiObject -query "SELECT * from Win32_DiskDrive" | Sort-Object
        foreach ( $drive in $drives ) {
            if ($drive.SerialNumber -ne $null) {
                [string]::Format( "{0}. {1}, Serial: {2}, Size: {3}GB",
                    $drive.DeviceID.substring(17), $drive.Model.Trim(),
                    $drive.SerialNumber.Trim(), [int]($drive.Size/1000000000) )
            } else {
                [string]::Format( "{0}. {1}, Size: {2}GB",
                    $drive.DeviceID.substring(17), $drive.Model.Trim(),
                    [int]($drive.Size/1000000000) )
            }
        }
        exit
    }

    if ($help) { usage }

    if (($util -lt 1) -or ($util -gt 100)) {
        "ERROR: Utilization must be between 1 and 100.`n"
        usage
    } else {
        $global:utilization = $util
    }

    if ( $outDir -eq "none" ) {
        $global:outDir = "${PWD}"
    } else {
        $global:outDir = "$outDir"
    }

    if ( $drive -ne "none" ) {
        $global:testMode = "cli"
        if ( $drive -notin (Get-Disk).Number ){
            Write-Error "The drive number `"$drive`" you entered does not exist.`n`n"
            usage
        }
        $global:physDrive = "\\.\PhysicalDrive$drive"
    } else {
        $global:testmode = "gui"
        $ok = IntroDialog
        if (-not ($ok) ) {
            exit
        }
    }

    $global:yes = $yes
    if ( -not $nullio ) {
        $global:ioengine = "windowsaio"
    } else {
        $global:ioengine = "null"
    }
    $global:quickie = $quickie
    $global:fastPrecond = $fastprecond

    # Do a sanity check that the selected drive does not show as a local drive letter
    Get-WMIObject Win32_LogicalDisk | Foreach-Object {
        $did = (Get-WmiObject -Query "Associators of {Win32_LogicalDisk.DeviceID='$($_.DeviceID)'} WHERE ResultRole=Antecedent").Path
        $dl = $_.DeviceID
        if ($did.RelativePath) {
            $part = $did.RelativePath.Split('"')[1]
            $pd = $part.split(',')[0].split('#')[1]
            if ($global:physDrive.ToLower() -eq "\\.\physicaldrive$pd") {
                if ($global:testmode -eq "cli") {
                    Write-Error "ERROR! Drive '$global:physdrive' is mounted as drive '$dl'!"
                    Write-Error "Aborting run, cannot run on mounted filesystem."
                    exit
                } else {
                    [System.Windows.Forms.MessageBox]::Show( "ERROR! Drive '$global:physdrive' is mounted as drive '$dl'!`nAborting run, cannot run on mounted filesystem.", "Fatal Error", 0, 48 ) | Out-Null
                    exit
                }
            }
        }
    }

}


function CollectSystemInfo()
{
    # Collect some OS and CPU information.

    # May want to put a window up while this happens.  GWMI is very slow

    $procs = [array](Get-WmiObject -class win32_processor) # Single-socket gives object, so coerce into array to match multisocket
    $global:cpu = $procs[0].Name.Trim()
    $cpuCount = ($procs[0].NumberOfCores).Count
    $cpuCores = ($procs[0] | Where DeviceID -eq "CPU0" ).NumberOfLogicalProcessors
    $global:cpuCores = $cpuCores * $cpuCount
    $global:cpuFreqMHz = ($procs[0] | Where DeviceID -eq "CPU0").MaxClockSpeed
    $os = Get-WmiObject Win32_OperatingSystem
    $global:uname = $os.Caption.Trim() + " - Build " + $os.BuildNumber.Trim() + " - ServicePack " + $os.ServicePackMajorVersion + "." + $os.ServicePackMinorVersion

    # Check if we're running in high-performance mode
    $plan = Get-WmiObject -Class win32_powerplan -Namespace root\cimv2\power -Filter "IsActive=True"
    if (-not ($plan.InstanceID -like "*8c5e7fda-e8bf-4a96-9a85-a6e23a8c635c*")) {
        function SetHighPerformance() {
            "Setting High Performance power scheme via POWERCFG"
            powercfg /setactive "8c5e7fda-e8bf-4a96-9a85-a6e23a8c635c"
        }
        if ($global:yes) {
            SetHighPerformance
        } elseif ($testmode -ne "gui") {
            "-" * 75
            "Power mode is not currently set to High Performance."
            "This may result in lowered test results."
            $cont = Read-Host "Would you like to enable this power setting now? (y/n)"
            "-" * 75
            if ($cont -ne "") {
                if ($cont.Substring(0, 1).ToLower() -eq "y" ) {
                    SetHighPerformance
                }
            }
        } else {
            $ret = [System.Windows.Forms.MessageBox]::Show(
                "System power mode not set to High Performance.`nThis may result in lowered test results.`nWould you like to enable High Performance Mode now?",
                "Verify Performance Mode", 4, 32)
            if ($ret -eq"yes" ) {
                SetHighPerformance
            }
        }
    }
}


function VerifyContinue()
{
    # User's last chance to abort the test.  Exit if they don't agree.
    if ( -not $global:yes ) {
        if ($testMode -ne "gui") { # text-mode prompt since we're running with command line options
            "-" * 75
            "WARNING! " * 9
            "THIS TEST WILL DESTROY ANY DATA AND FILESYSTEMS ON $global:physDrive"
            $cont = Read-Host "Please type the word `"yes`" and hit return to continue, or anything else to abort"
            "-" * 75
            if ( $cont -ne "yes" ) {
                "Performance test aborted, drive is untouched."
                exit
            }
        } else {
            # Do it in a messagebox since we're running GUI-wise
            $ret = [System.Windows.Forms.MessageBox]::Show(
                "Test selected to run on $global:physDrive.`nALL DATA WILL BE ERASED ON THIS DRIVE`nContinue with testing?",
                "Verify the device to test", 4)
            if ($ret -ne "yes" ) {
                "Performance test aborted, drive is untouched."
                exit
            }
        }
    }
}

function CollectDriveInfo()
{
    # Get important device information, exit if not possible.
    # We absolutely need this information
    $global:physDriveBase = ([io.fileinfo]$global:physDrive).BaseName
    $global:physDriveNo = $global:physDrive.Substring(17)
    $global:physDriveBytes=(GET-WMIOBJECT win32_diskdrive | where DeviceID -eq $global:physDrive).Size
    $global:physDriveGB=[long]($global:physDriveBytes/(1000*1000*1000))
    $global:physDriveGiB=[long]($global:physDriveBytes/(1024*1024*1024))
    $global:testcapacity = [long](($global:physDriveGiB * $global:utilization) / 100)
    # This is just nice to have
    $drive = (Get-Disk | Where-Object { $_.Number -eq $global:physDriveNo })
    $global:model = $drive.Model.ToString().Trim()
    if ($drive.SerialNumber -ne $null ) { $global:serial = $drive.SerialNumber.ToString().Trim() }
    else { $global:serial = "UNKNOWN" }
}


# Set up names for all output/input files, place headers on CSVs.
function CSVInfoHeader {
    if ($global:fastPrecond -eq $false) { $prefix = "" }
    else { $prefix = "FASTPRECOND-" }

    #Headers to the CSV file (ending up in the ODS at the test end)
    "Drive,$prefix$global:physDrive"
    "Model,$prefix$global:model"
    "Serial,$prefix$global:serial"
    "AvailCapacity,$prefix$global:physDriveGiB,GiB"
    "TestedCapacity,$prefix$global:testcapacity,GiB"
    "CPU,$prefix$global:cpu"
    "Cores,$prefix$global:cpuCores"
    "Frequency,$prefix$global:cpuFreqMHz"
    "OS,$prefix$global:uname"
    "FIOVersion,$prefix$global:fioVerString"
}

function SetupFiles()
{
    # Datestamp for run output files
    $global:ds=(Get-Date).ToString("yyyy-MM-dd_HH-mm-ss")

    # The unique suffix we generate for all output files
    $suffix="${global:physDriveGB}GB_${global:cpuCores}cores_${global:cpuFreqMHz}MHz_${global:physDriveBase}_${env:computername}_${global:ds}"

    # Need to worry about normalizing passed in directory names, or else non-absolute output paths will resolve to c:\windows\system32\...
    if ( -not ( Test-Path -Path $global:outDir ) ) {
        # New-Item used PWD, so we're OK here
        New-Item -ItemType directory -Path $global:outDir | Out-Null
    }
    # Now resolve to c:\... path and put back to global for sanity.
    $global:outDir = Resolve-Path $global:outDir

    # The "details" directory contains the raw output of each FIO run
    $global:details = "${global:outDir}\details_${suffix}"
    # The "details" directory contains the raw output of each FIO run
    if ( Test-Path -Path $global:details ) {
        Remove-Item -Recurse -Force $global:details | Out-Null
    }
    New-Item -ItemType directory -Path $global:details | Out-Null

    # Copy this script into it for posterity
    Copy-Item $scriptName $global:details

    # Files we're going to generate, encode some system info in the names
    # If the output files already exist, erase them
    $global:testcsv = "${global:details}\ezfio_tests_${suffix}.csv"
    if (Test-Path $global:testcsv) { Remove-Item $global:testcsv }
    CSVInfoHeader > $global:testcsv
    "Type,Write %,Block Size,Threads,Queue Depth/Thread,IOPS,Bandwidth (MB/s),Read Latency (us),Write Latency (us)" >> $global:testcsv
    $global:timeseriescsv ="${global:details}\ezfio_timeseries_${suffix}.csv"
    $global:timeseriesclatcsv ="${global:details}\ezfio_timeseriesclat_${suffix}.csv"
    $global:timeseriesslatcsv ="${global:details}\ezfio_timeseriesslat_${suffix}.csv"
    if (Test-Path $global:timeseriescsv) { Remove-Item $global:timeseriescsv }
    if (Test-Path $global:timeseriesclatcsv) { Remove-Item $global:timeseriesclatcsv }
    if (Test-Path $global:timeseriesslatcsv) { Remove-Item $global:timeseriesslatcsv }
    CSVInfoHeader > $global:timeseriescsv
    CSVInfoHeader > $global:timeseriesclatcsv
    CSVInfoHeader > $global:timeseriesslatcsv
    "IOPS" >> $global:timeseriescsv  # Add IOPS header
    "CLAT-read,CLAT-write" >> $global:timeseriesclatcsv
    "SLAT-read,SLAT-write" >> $global:timeseriesslatcsv

    # ODS input and output files
    $global:odssrc = "${PWD}\original.ods"
    $global:odsdest = "${global:outDir}\ezfio_results_${suffix}.ods"
    if (Test-Path $global:odsdest) { Remove-Item $global:odsdest }
}


function TestName ($seqrand, $wmix, $bs, $threads, $iodepth)
{
    # Return full path and filename prefix for test of specified params
    $testfile  = $global:details + "\Test" + $seqrand + "_w" + [string]$wmix
    $testfile += "_bs" + [string]$bs + "_threads" + [string]$threads + "_iodepth"
    $testfile += [string]$iodepth + "_" + $global:physDriveBase + ".out"
    return $testfile
}

# The actual functions that run FIO, in a string so that we can do a Start-Job using it.
$global:jobutils = @'

function TestName ($seqrand, $wmix, $bs, $threads, $iodepth)
{
    # Return full path and filename prefix for test of specified params
    $testfile  = $details + "\Test" + $seqrand + "_w" + [string]$wmix
    $testfile += "_bs" + [string]$bs + "_threads" + [string]$threads + "_iodepth"
    $testfile += [string]$iodepth + "_" + $physDriveBase + ".out"
    return $testfile
}

function SequentialConditioning
{
    # Sequentially fill the complete capacity of the drive once.
    # Note that we can't use regular test runner because this test needs
    # to run for a specified # of bytes, not a specified # of seconds.
    if ( $quickie ) {
        $size = "1G"
    } else {
        $size = "${testcapacity}G"
    }
    . $fio "--name=SeqCond" "--readwrite=write" "--bs=128k" "--ioengine=$ioengine" "--iodepth=64" "--direct=1" "--filename=$physDrive" "--size=$size" "--thread" | Out-Null
    if ( $LastExitCode -ne 0 ) {
        Write-Output "ERROR" "ERROR" "ERROR"
    } else {
        Write-Output "DONE" "DONE" "DONE"
    }
}

function RandomConditioning
{
    # Randomly write entire device for the full capacity
    # Note that we can't use regular test runner because this test needs
    # to run for a specified # of bytes, not a specified # of seconds.
    if ( $quickie ) {
        $size = "1G"
    } else {
        $size = "${testcapacity}G"
    }
    . $fio "--name=RandCond" "--readwrite=randwrite" "--bs=4k" "--invalidate=1" "--end_fsync=0" "--group_reporting" "--direct=1" "--filename=$physDrive"  "--size=$size" "--ioengine=$ioengine" "--iodepth=256" "--norandommap" "--randrepeat=0" "--thread" | Out-Null
    if ( $LastExitCode -ne 0 ) {
        Write-Output "ERROR" "ERROR" "ERROR"
    } else {
        Write-Output "DONE" "DONE" "DONE"
    }
}

# Taken from fio_latency2csv.py
function plat_idx_to_val( $idx, $FIO_IO_U_PLAT_BITS, $FIO_IO_U_PLAT_VAL )
{
    # MSB <= (FIO_IO_U_PLAT_BITS-1), cannot be rounded off. Use
    # all bits of the sample as index
    if ($idx -lt ($FIO_IO_U_PLAT_VAL -shl 1)) {
        return $idx
    }
    # Find the group and compute the minimum value of that group
    $error_bits = ($idx -shr $FIO_IO_U_PLAT_BITS) - 1
    $base = 1 -shl ($error_bits + $FIO_IO_U_PLAT_BITS)
    # Find its bucket number of the group
    $k = $idx % $FIO_IO_U_PLAT_VAL
    # Return the mean of the range of the bucket
    return ($base + (($k + 0.5) * (1 -shl $error_bits)))
}

function WriteExceedance($j, $rdwr, $outfile)
{
    # Generate an exceedance CSV for read or write from JSON output.
    if ($fioOutputFormat -eq "json") {
        return # This data not present in JSON format, only JSON+
    }
    $ios = $j.jobs[0].$rdwr.total_ios
    if ( $ios -gt 0 ) {
        $runttl = 0;
        # FIO 2.99 changed this to use saner latency bucketing, no semi-log needed
        if ($j.jobs[0].$rdwr.clat_ns) {
            # This is very inefficient, but need to convert from object.property's to sorted ints...
            $lat_ns = @()
            foreach ($n in ((Get-Member -inputObject $j.jobs[0].$rdwr.clat_ns.bins -MemberType Properties).name) ) {
                $lat_ns += [long]$n
            }
            foreach ($b in ($lat_ns | sort-object)) {
                $lat_us = [float]($b) / 1000.0
                $cnt = [int]$j.jobs[0].$rdwr.clat_ns.bins.$b
                $runttl += $cnt
                $pctile = 1.0 - [float]$runttl / [float]$ios;
                if ( $cnt -gt 0 ) {
                    "$lat_us,$pctile" >> $outfile
                }
            }
        } else {
            $plat_bits = $j.jobs[0].$rdwr.clat.bins.FIO_IO_U_PLAT_BITS
            $plat_val = $j.jobs[0].$rdwr.clat.bins.FIO_IO_U_PLAT_VAL
            foreach ($b in 0..[int]$j.jobs[0].$rdwr.clat.bins.FIO_IO_U_PLAT_NR) {
                $cnt = [int]$j.jobs[0].$rdwr.clat.bins.$b
                $runttl += $cnt
                $pctile = 1.0 - [float]$runttl / [float]$ios
                if ( $cnt -gt 0 ) {
                    $p2idx = plat_idx_to_val $b $plat_bits $plat_val
                    "${p2idx},${pctile}" >> $outfile
                }
            }
        }
    }
}

function CombineThreadOutputs($suffix, $outcsv, $lat, $runtime, $extra_runtime)
{
    # Merge all FIO iops/lat logs across all servers"""
    # The lists may be called "iops" but the same works for clat/slat
    $testtime = $runtime + $extra_runtime
    $iops = New-Object 'float[]' $testtime
    # For latencies, need to keep the _w and _r separate
    $iops_w = New-Object 'float[]' $testtime
    $filecnt = 0
    $fileglob = "$testfile$suffix.*log"
    Get-ChildItem $fileglob | ForEach-Object {
        $filename = $_.FullName
        $filecnt++
    $csvhdr = 'timestamp', 'value', 'wr', 'ign'
    $lines = Import-Csv -Path $filename -Header $csvhdr
    $lineidx = 0
        # Set time 0 IOPS to first values
        $riops = [float]0.0
        $wiops = [float]0.0
        $nexttime = [float]0.0
        for ($x=0; $x -lt $testtime; $x++) {
            if ( -not $lat ) {
                $iops[$x] = [float]$iops[$x] + [float]$riops + [float]$wiops
            } else {
                $iops[$x] = [float]$iops[$x] + [float]$riops
                $iops_w[$x] = [float]$iops_w[$x] + [float]$wiops
            }
            while (($lineidx -lt $lines.Count) -and ($nexttime -lt $x)) {
                $nexttime = $lines[$lineidx].timestamp / 1000.0
                if ( $lines[$lineidx].wr -eq 1 ) {
                    $wiops = [int]$lines[$lineidx].value
                } else {
                    $riops = [int]$lines[$lineidx].value
                }
                $lineidx++
            }
        }
    }

    # Generate the combined CSV
    for ($x=[int]($extra_runtime / 2); $x -lt ($runtime + $extra_runtime); $x++) {
        if ( $lat ) {
            $a = [float]$iops[$x] / [float]$filecnt
            $b = [float]$iops_w[$x] / [float]$filecnt
            "{0:f1},{1:f1}" -f $a, $b >> $outcsv
        } else {
        $a = $iops[$x]
            "{0:f0}" -f $a >> $outcsv
        }
    }
}


function RunTest
{
    # Runs the specified test, generates output CSV lines.

    # Output file names
    $testfile = TestName $seqrand $wmix $bs $threads $iodepth

    if ( $seqrand -eq "Seq" ) { $rw = "rw" }
    else { $rw = "randrw" }

    if ( $iops_log ) {
        $extra_runtime = 10
    } else {
        $extra_runtime = 0
    }
    $testtime = $runtime + $extra_runtime

    $cmd  = ("--name=test", "--readwrite=$rw", "--rwmixwrite=$wmix")
    $cmd += ("--bs=$bs", "--invalidate=1", "--end_fsync=0")
    $cmd += ("--group_reporting", "--direct=1", "--filename=$physDrive")
    $cmd += ("--size=${testcapacity}G", "--time_based", "--runtime=$testtime")
    $cmd += ("--ioengine=$ioengine", "--numjobs=$threads")
    $cmd += ("--iodepth=$iodepth", "--norandommap", "--randrepeat=0")
    if ( $iops_log ) {
        $cmd += ("--write_iops_log=$testfile")
        $cmd += ("--write_lat_log=$testfile")
        $cmd += ("--log_avg_msec=1000")
        $cmd += ("--log_unix_epoch=0")
    }
    $cmd += ("--thread", "--output-format=$fioOutputFormat", "--exitall")
    $fio + " " + [string]::Join(" ", $cmd) | Out-File $testfile

    # Check that the IO size is usable.  Some SSDs are only 4K logical sectors
    $minblock = (Get-Disk | Where-Object { $_.Number -eq $global:physDriveNo }).LogicalSectorSize
    if ( $bs -lt $minblock ) {
        "Test not run because block size $bs below minimum size $minblock" | Out-File -Append $testfile
        "3;" + "0;" * 100 | Out-File -Append $testfile # Bogus 0-filled result line
        "1,1" | Out-File "${testfile}.exc.read.csv"
        "1,1" | Out-File "${testfile}.exc.write.csv"
        "$seqrand,$wmix,$bs,$threads,$iodepth,0,0,0,0" | Out-File -Append $testcsv
        Write-Output "SKIP" "SKIP" "SKIP"
        return
    }

    . $fio @cmd | Out-File -Append $testfile

    if ( $LastExitCode -ne 0 ) {
        Write-Output "ERROR" "ERROR" "ERROR"
        return # Don't process this one, it was error'd out!
    }

    if ( $iops_log ) {
        CombineThreadOutputs '_iops' $timeseriescsv $false $runtime $extra_runtime
        CombineThreadOutputs '_clat' $timeseriesclatcsv $true $runtime $extra_runtime
        CombineThreadOutputs '_slat' $timeseriesslatcsv $true $runtime $extra_runtime
    }

    # Thanks to @BryanTuttle.  Skip any FIO output before the JSON open-bracket
    $LineSkip=0
    foreach ($line in Get-Content $testfile) {
	    if ($line -match '^{') { break }
		else {$LineSkip++}
    }

    $j = ConvertFrom-Json "$(Get-Content $testfile | select -Skip $LineSkip)"
    $rdiops = [float]($j.jobs[0].read.iops);
    $wriops = [float]($j.jobs[0].write.iops);
    $rlat = [float]($j.jobs[0].read.lat_ns.mean) / 1000.0;
    if ($rlat -le 0.0001) { $rlat = [float]($j.jobs[0].read.lat.mean); }
    $wlat = [float]($j.jobs[0].write.lat_ns.mean) / 1000.0;
    if ($wlat -le 0.0001) { $wlat = [float]($j.jobs[0].wlat.lat.mean); }
    $iops = "{0:F0}" -f ($rdiops + $wriops)
    # Locale output is not wanted here, manually make a decimal string.  Ugh
    $lat = "{0:F1}" -f ([math]::Max($rlat, $wlat))
    $mbpsfloat = (( ($rdiops+$wriops) * $bs ) / ( 1024.0 * 1024.0 ))
    "{0:f1}" -f $mbpsfloat | Set-Variable mbps
    $lat = "{0:F1}" -f ([math]::Max($rlat, $wlat)) # This is just displayed, use native locale
    "$seqrand,$wmix,$bs,$threads,$iodepth,$iops,$mbps,$rlat,$wlat" | Out-File -Append $testcsv

    WriteExceedance $j "read" "${testfile}.exc.read.csv"
    WriteExceedance $j "write" "${testfile}.exc.write.csv"

    Write-Output $iops $mbps $lat
}
'@


function DefineTests {
    # Generate the work list for the main worker into OC.

    # What we're shmoo-ing across
    $bslist = (512, 1024, 2048, 4096, 8192, 16384, 32768, 65536, 131072)
    $qdlist = (1, 2, 4, 8, 16, 32, 64, 128, 256)
    $threadslist = (1, 2, 4, 8, 16, 32, 64, 128, 256)
    $shorttime = 120 # Runtime of point tests
    $longtime = 1200 # Runtime of long-running tests
    if ( $quickie ) {
        $shorttime = [int]($shorttime / 10)
        $longtime = [int]($longtime / 10)
    }
    function AddTest( $name, $seqrand, $writepct, $blocksize, $threads, $qdperthread, $desc, $cmdline ) {
        if ($threads -eq "") { $qd = '' } else { $qd = ([int]$threads) * ([int]$qdperthread) }
        if ($blocksize -ne "") { if ($blocksize -lt 1024) { $bsstr = "${blocksize}b" } else { $bsstr = "{0:N0}K" -f ([int]$blocksize/1024) } }
        if ($writepct -ne "" ) { $writepct = [string]$writepct + "%" }
        $dat = New-Object psobject -Property @{ name=$name; seqrand=$seqrand; writepct=$writepct
            bs=$bsstr; qd = $qd; qdperthread = $qdperthread; bw = ''; iops= ''; lat = ''; desc = $desc;
            cmdline = $cmdline }
        $global:oc.Add( $dat )
    }

    function DoAddTest {
        AddTest $testname $seqrand $wmix $bs $threads $iodepth $desc "$global:globals; $global:jobutils; `$iops_log=$iops_log; `$seqrand=`"$seqrand`"; `$wmix=$wmix; `$bs=$bs; `$threads=$threads; `$iodepth=$iodepth; `$runtime=$runtime; RunTest"
    }

    function AddTestBSShmoo {
        AddTest $testname 'Preparation' '' '' '' '' '' "$global:globals; `"$testname`" >> `"$global:testcsv`"; Write-Output ' ' ' ' ' '"
        foreach ($bs in $bslist ) { $desc = "$testname, BS=$bs"; DoAddTest }
    }

    function AddTestQDShmoo {
        AddTest $testname 'Preparation' '' '' '' '' '' "$global:globals; `"$testname`" >> `"$global:testcsv`"; Write-Output ' ' ' ' ' '"
        foreach ($iodepth in $qdlist ) { $desc = "$testname, QD=$iodepth"; DoAddTest }
    }

    function AddTestThreadsShmoo {
        AddTest $testname 'Preparation' '' '' '' '' '' "$global:globals; `"$testname`" >> `"$global:testcsv`"; Write-Output ' ' ' ' ' '"
        foreach ($threads in $threadslist) { $desc = "$testname, Threads=$threads"; DoAddTest }
    }

    AddTest 'Sequential Preconditioning' 'Seq Pass 1' '100' '131072' '1' '256' 'Sequential Preconditioning' "$global:globals; $global:jobutils; SequentialConditioning;"
    if ($global:fastPrecond -ne $true) {
	    AddTest 'Sequential Preconditioning' 'Seq Pass 2' '100' '131072' '1' '256' 'Sequential Preconditioning' "$global:globals; $global:jobutils; SequentialConditioning;"
    }

    $testname = "Sustained Multi-Threaded Sequential Read Tests by Block Size"
    $seqrand = "Seq"; $wmix=0; $threads=1; $runtime=$shorttime; $iops_log="`$false"; $iodepth=256
    AddTestBSShmoo

    $testname = "Sustained Multi-Threaded Random Read Tests by Block Size"
    $seqrand = "Rand"; $wmix=0; $threads=16; $runtime=$shorttime; $iops_log="`$false"; $iodepth=16
    AddTestBSShmoo

    $testname = "Sequential Write Tests with Queue Depth=1 by Block Size"
    $seqrand = "Seq"; $wmix=100; $threads=1; $runtime=$shorttime; $iops_log="`$false"; $iodepth=1
    AddTestBSShmoo

    if ($global:fastPrecond -ne $true) {
        AddTest 'Random Preconditioning' 'Rand Pass 1' '100' '4096' '1' '256' 'Random Preconditioning' "$global:globals; $global:jobutils; RandomConditioning;"
        AddTest 'Random Preconditioning' 'Rand Pass 2' '100' '4096' '1' '256' 'Random Preconditioning' "$global:globals; $global:jobutils; RandomConditioning;"
    }

    $testname = "Sustained 4KB Random Read Tests by Number of Threads"
    $seqrand = "Rand"; $wmix=0; $bs=4096; $runtime=$shorttime; $iops_log="`$false"; $iodepth=1
    AddTestThreadsShmoo

    $testname = "Sustained 4KB Random mixed 30% Write Tests by Number Threads"
    $seqrand = "Rand"; $wmix=30; $bs=4096; $runtime=$shorttime; $iops_log="`$false"; $iodepth=1
    AddTestThreadsShmoo

    $testname = "Sustained Perf Stability Test - 4KB Random 30% Write for 20 minutes"
    $desc = $testname
    AddTest $testname 'Preparation' '' '' '' '' '' "$global:globals; `"$testname`" >> `"$global:testcsv`"; Write-Output ' ' ' ' ' '"
    $seqrand = "Rand"; $wmix=30; $bs=4096; $runtime=$longtime; $iops_log="`$true"; $iodepth=1; $threads=256
    DoAddTest

    $testname = "Sustained 4KB Random Write Tests by Number of Threads"
    $seqrand = "Rand"; $wmix=100; $bs=4096; $runtime=$shorttime; $iops_log="`$false"; $iodepth=1
    AddTestThreadsShmoo

    $testname = "Sustained Multi-Threaded Random Write Tests by Block Size"
    $seqrand = "Rand"; $wmix=100; $runtime=$shorttime; $iops_log="`$false"; $iodepth=16; $threads=16
    AddTestBSShmoo
}


function RunAllTests()
{
    # Iterate through the OC work queue and run each job, show progress.

    function UpdateView {
        # Updates the grid to reflect new data, scrolls to selection
        $t_testList.ItemsSource.Refresh()
        $t_testList.UpdateLayout()
        $t_testList.ScrollIntoView($t_testList.SelectedItem)
    }

    function NotifyIcon {
        # NotifyIcon needs to run as separate Powerhell process because
        # a WPF form will block other events (like the notify-clicked) until
        # it returns control to PowerShell

        # Pass destination into the block through the child's environment
        [System.Environment]::SetEnvironmentVariable("ods", $global:odsdest)
        $proc = Start-Process -PassThru (Get-Command powershell.exe) -WindowStyle Hidden -ArgumentList ( "-Command", {
            Add-Type -AssemblyName PresentationFramework, System.Windows.Forms
            echo ([System.Environment]::GetEnvironmentVariable("'ods'"))
            # Add a NotifyIcon that, when clicked, will open the results spreadsheet
            $global:notify = New-Object System.Windows.Forms.NotifyIcon
            $global:notify.Icon = [System.Drawing.SystemIcons]::Information
            $global:notify.BalloonTipIcon = "'Info'"
            $global:notify.BalloonTipText = "'The ezFIO test series has completed and result spreadsheet may be opened.'"
            $global:notify.Text = "'Click to open the ezFIOresult spreadsheet'"
            $global:notify.BalloonTipTitle = "'ezFIO Test Completion'"
            $global:notify.Visible = $True
            # Using the add_BalloonTipClicked() seemed to fault every time
            Unregister-Event -SourceIdentifier click_event -ErrorAction SilentlyContinue
            Register-ObjectEvent $notify Click -sourceIdentifier click_event -Action {
                Invoke-Item ([System.Environment]::GetEnvironmentVariable("'ods'"))
                $global:notify.Dispose()
                $global:notify = $null
                } | Out-Null
            Unregister-Event -SourceIdentifier balloonclick_event -ErrorAction SilentlyContinue
            Register-ObjectEvent $notify BalloonTipClicked -SourceIdentifier balloonclick_event -Action {
                Invoke-Item ([System.Environment]::GetEnvironmentVariable("'ods'"))
                $global:notify.Dispose()
                $global:notify = $null
            } | Out-Null
            $notify.ShowBalloonTip(10000)
            while ( $global:notify -ne $null ) { sleep 1 }
        } )
        return $proc
    }

    $xaml = @'
<Window x:Class="Window3"
    xmlns="http://schemas.microsoft.com/winfx/2006/xaml/presentation"
    xmlns:x="http://schemas.microsoft.com/winfx/2006/xaml"
    Title="ezFIO Test Progress" Height="562.084" Width="682.007">
    <Window.Resources>
        <Style x:Key="CellRightAlign">
            <Setter Property="Control.HorizontalAlignment" Value="Right" />
        </Style>
        <Style x:Key="CellCenterAlign">
            <Setter Property="Control.HorizontalAlignment" Value="Center" />
        </Style>
    </Window.Resources>
    <Grid>
        <Label Content="Testing Drive:" HorizontalAlignment="Left" Margin="83,23,0,0" VerticalAlignment="Top"/>
        <Label x:Name="testingDrive" Content="\\physicaldrive0" HorizontalAlignment="Left" Margin="167,23,0,0" VerticalAlignment="Top"/>
        <Label Content="Current Test Runtime:" HorizontalAlignment="Left" Margin="40,75,0,0" VerticalAlignment="Top"/>
        <Label x:Name="testRuntime" Content="00:00:00" HorizontalAlignment="Left" Margin="167,75,0,0" VerticalAlignment="Top"/>
        <Label Content="Current Test:" HorizontalAlignment="Left" Margin="88,49,0,0" VerticalAlignment="Top"/>
        <Label x:Name="currentTest" Content="BS 4K, QD 32, WR 100%" HorizontalAlignment="Left" Margin="167,49,0,0" VerticalAlignment="Top"/>
        <DataGrid x:Name="testList" HorizontalAlignment="Left" Margin="36,146,0,0" VerticalAlignment="Top" Height="321" Width="600" GridLinesVisibility="None" HeadersVisibility="Column">
            <DataGrid.GroupStyle>
                <GroupStyle>
                    <GroupStyle.HeaderTemplate>
                        <DataTemplate>
                            <TextBlock Text="{Binding Items[0].name}"/>
                        </DataTemplate>
                    </GroupStyle.HeaderTemplate>
                </GroupStyle>
            </DataGrid.GroupStyle>
            <DataGrid.Columns>
                <DataGridTextColumn Header="Access Pattern" Binding="{Binding seqrand}" CanUserSort="false" CanUserReorder="false" IsReadOnly="true"/>
                <DataGridTextColumn Header="Write %" Binding="{Binding writepct}" CanUserSort="false" CanUserReorder="false" IsReadOnly="true" ElementStyle="{StaticResource CellRightAlign}"/>
                <DataGridTextColumn Header="Block Size" Binding="{Binding bs}" CanUserSort="false" CanUserReorder="false" IsReadOnly="true" ElementStyle="{StaticResource CellRightAlign}"/>
                <DataGridTextColumn Header="Queue Depth" Binding="{Binding qd}" CanUserSort="false" CanUserReorder="false" IsReadOnly="true" ElementStyle="{StaticResource CellRightAlign}"/>
                <DataGridTextColumn Header="        " Binding="{Binding blank}" CanUserSort="false" CanUserReorder="false" IsReadOnly="true" ElementStyle="{StaticResource CellRightAlign}"/>
                <DataGridTextColumn Header="IOPS" Binding="{Binding iops}" CanUserSort="false" CanUserReorder="false" IsReadOnly="true" ElementStyle="{StaticResource CellRightAlign}"/>
                <DataGridTextColumn Header="Bandwidth (MB/s)" Binding="{Binding bw}" CanUserSort="false" CanUserReorder="false" IsReadOnly="true" ElementStyle="{StaticResource CellRightAlign}"/>
                <DataGridTextColumn Header="Latency (us)" Binding="{Binding lat}" CanUserSort="false" CanUserReorder="false" IsReadOnly="true" ElementStyle="{StaticResource CellRightAlign}"/>
            </DataGrid.Columns>
        </DataGrid>
        <Button x:Name="openSpreadsheet" Content="Open Graphs Spreadsheet" HorizontalAlignment="Left" Margin="236,486,0,0" Width="200" Height="27" VerticalAlignment="Top"/>
        <Label Content="Total Test Runtime:" HorizontalAlignment="Left" Margin="53,102,0,0" VerticalAlignment="Top"/>
        <Label x:Name="totalRuntime" Content="00:00:00" HorizontalAlignment="Left" Margin="167,102,0,0" VerticalAlignment="Top"/>
    </Grid>
</Window>
'@


    # The test window
    $t = WindowFromXAML $xaml 't'
    $t.Icon = $global:iconBitmap

    $t.add_Loaded( {
        $t.Activate()
        $t_testList.Focus()

        $global:step = -1 # Which test we're on
        $global:curjob = $null # Which process is running
        $global:totalStarttime = Get-Date

        # The NotifyIcon process info
        $global:notifyProc = $null

        $t_testingDrive.Content = [string]::Format("{0}, {1}({2}), {3}GB", $global:physDrive, $global:model, $global:serial, $global:testcapacity )
        $t_openSpreadsheet.IsEnabled = $false
        $t_currentTest.Content = "Starting up..."

        $t_testList.CanUserAddRows = $false
        $t_testList.AutoGenerateColumns = $false
        $t_testList.ItemsSource = $null

        $lview = [System.Windows.Data.ListCollectionView]$global:oc
        $lview.GroupDescriptions.Add((new-object System.Windows.Data.PropertyGroupDescription "name"))
        $t_testList.ItemsSource = $lview

        # Poor man's threading/event driven
        $global:timer = new-object System.Windows.Threading.DispatcherTimer
        $global:timer.Interval = [TimeSpan]"0:0:1.00"
        $global:timer.Add_Tick(
        {
            # If there's a running job, update the runtime if not done, and capture the results if finished
            if ($global:curjob -ne $null)
            {
                if ( $global:curjob.State -match 'running' )
                {
                    $now = Get-Date
                    $delta = $now - $global:starttime
                    $ts = [timespan]::FromTicks($delta.Ticks)
                    $t_testRuntime.Content = $ts.ToString("hh\:mm\:ss")
                    $delta = $now - $global:totalstarttime
                    $ts = [timespan]::FromTicks($delta.Ticks)
                    $t_totalRuntime.Content = $ts.ToString("hh\:mm\:ss")
                } else {
                    # Job just finished, let's read out answers
                    $q = Receive-Job $global:curjob
                    $global:oc[$global:step].iops = $q[0]
                    $global:oc[$global:step].bw = $q[1]
                    $global:oc[$global:step].lat = $q[2]
                    $ign = 0.0
                    if ([float]::TryParse($global:oc[$global:step].iops, [ref]$ign)) { $global:oc[$global:step].iops = [string]::Format("{0:N0}", [float]$global:oc[$global:step].iops) }
                    if ([float]::TryParse($global:oc[$global:step].bw, [ref]$ign)) { $global:oc[$global:step].bw = [string]::Format("{0:N1}", [float]$global:oc[$global:step].bw) }
                    $t_testList.SelectedIndex = $global:step
                    UpdateView
                    if ($global:oc[$global:step].iops -eq "ERROR") {
                        $global:step = 9999 # Skip all the other tests
                        [System.Windows.Forms.MessageBox]::Show( "ERROR!  FIO job did not complete successfully.  Aborting further runs.", "Fatal Error", 0, 48 ) | Out-Null
                    }
                }
            }
            # If there's no running job (last one finished), start a new one
            if ( ($global:curjob -eq $null) -or ( -not ($global:curjob.State -match 'running' ) ) ){
                $global:step = $global:step + 1
                if ($global:step -lt $t_testList.Items.Count) {
                    $t_testList.SelectedIndex = $global:step
                    $global:cmdline = $t_testList.Items[$global:step].cmdline
                    $t_currentTest.Content = $t_testList.Items[$global:step].desc
                    # Powershell won't have the $globals in the Start-Job context, so expand here
                    $fullcmd = "Start-Job { $global:cmdline }"
                    $global:curjob = Invoke-Expression $fullcmd
                    $global:starttime = Get-Date
                    $global:oc[$global:step].iops = "Running"
                    $global:oc[$global:step].bw = "Running"
                    $global:oc[$global:step].lat = "Running"
                    $t_testList.SelectedIndex = $global:step
                    UpdateView
                } else {
                    $global:timer.Stop()
                    $global:curjob = $null
                    $t_testList.SelectedIndex = $null
                    if ($global:step -lt 9999) {
                        $t_currentTest.Content = "Completed"
                        GenerateResultODS
                        $t_openSpreadsheet.IsEnabled = $true
                        $global:notifyProc = NotifyIcon
                    } else {
                        $t_currentTest.Content = "ERROR"
                    }
                }
            }
        } )
        $global:timer.Start()
    } )

    $t.add_Closing( {
        $global:timer.Stop()
    } )

    $t_openSpreadsheet.add_Click( {
        # Just open the file using default application
        Invoke-Item $global:odsdest
        $t.close()
    } )

    $t.ShowDialog() | Out-Null

    # Clean up the notifyicon process
    if ($global:notifyProc -ne $null) {
        if (-not ($global:notifyProc.HasExited)) { $global:notifyProc.Kill() }
    }
}


function RunAllTestsCLI()
{
    # CLI mode will short-circuit, run much simpler path and output only text

    # Determine some column widths to make format specifiers for CLI mode outputs
    $maxlen = 0
    foreach ($o in $global:oc) {
        $maxlen = [math]::max($maxlen, $o.desc.length)
    }
    $descfmt = "{0,-" + [string]$maxlen + "}"
    $resfmt = "{1,8} {2,9} {3,8}"
    $fmtstr = $descfmt + " " + $resfmt

    "*" * [string]::format( $fmtstr, "", "", "", "").length
    "ezFio test parameters:"

    $fmtinfo="{0,-20}: {1}"
    [string]::format( $fmtinfo, "Drive", $global:physDrive )
    [string]::format( $fmtinfo, "Model", $global:model )
    [string]::format( $fmtinfo, "Serial", $global:serial )
    [string]::format( $fmtinfo, "AvailCapacity", [string]$global:physDriveGiB + " GiB")
    [string]::format( $fmtinfo, "TestedCapacity", [string]$global:testcapacity + " GiB")
    [string]::format( $fmtinfo, "CPU", $global:cpu)
    [string]::format( $fmtinfo, "Cores", $global:cpuCores)
    [string]::format( $fmtinfo, "Frequency", $global:cpuFreqMHz)
    [string]::format( $fmtinfo, "FIO Version", $global:fioVerString)

    ""
    [string]::format( $fmtstr, "Test Description", "BW(MB/s)", "IOPS", "Lat(us)")
    [string]::format( $fmtstr, "-"*$maxlen, "-"*8, "-"*9, "-"*8)

    foreach ($o in $global:oc) {
        if ( $o.desc -eq "" ) {
            # This is a header-printing job, don't thread out
            [string]::format( $fmtstr, "---" + $o.name + "---", "", "", "")
            [Console]::Out.Flush()
            Invoke-Expression $o.cmdline | Out-Null
        } else {
            # This is a real test job, print some stuff, execute, then print output
            Write-Host -NoNewline ([string]::format($descfmt, $o.desc))
            [Console]::Out.Flush()
            $q = Invoke-Expression $o.cmdline
            $iops = $q[0]
            $mbps = $q[1]
            $lat = $q[2]
            Write-Host ([string]::format($resfmt, "", $mbps, $iops, $lat))
            if ($mbps -eq "ERROR") {
                "ERROR!  FIO job did not complete successfully.  Aborting further runs."
                return
            }
            [Console]::Out.Flush()
        }
    }
    GenerateResultODS
    "`nCOMPLETED!  Output file: $global:odsdest"
    return
}


function GenerateResultODS()
{
    # Builds a new ODS spreadsheet w/graphs from generated test CSV files.

    function GetContentXMLFromODS( $odssrc )
    {
        # Extract content.xml from an ODS file, where the sheet lives.
        $ziparchive = [System.IO.Compression.ZipFile]::Open( $odssrc, [System.IO.Compression.ZipArchiveMode]::Read )
        $zipentry = $ziparchive.GetEntry("content.xml")
        $reader = New-Object System.IO.StreamReader( $zipentry.Open() )
        $contentobj = $reader.ReadToEnd()
        $reader.Close()
        $ziparchive.Dispose()
        return $contentobj -replace "`n","" -replace "`r",""
    }

    function ReplaceSheetWithCSV_regex($sheetName, $csvName, $xmltext)
    {
        # Replace a named sheet with the contents of a CSV file.
        $newt = "<table:table table:name="
        $newt = $newt + "`"$sheetName`"" + ' table:style-name="ta1" > <table:table-column table:style-name="co1" table:default-cell-style-name="Default"/>'
        Get-Content $csvName | ForEach-Object {
            $newt = $newt + "<table:table-row table:style-name=`"ro1`">"
            foreach ($val in ($_.Split(','))) {
                $dbl = 0.0
                if ( [System.Double]::TryParse( $val, [ref]$dbl ) ) {
                    $newt = $newt + "<table:table-cell office:value-type=`"float`" office:value=`"$val`"><text:p>$val</text:p></table:table-cell>"
                } else {
                    $newt = $newt + "<table:table-cell office:value-type=`"string`"><text:p>$val</text:p></table:table-cell>"
                }
            }
            $newt = $newt + "</table:table-row>"
        }
        $newt = $newt + "</table:table>"
        $searchstr = "<table:table table:name=`"$sheetName`".*?</table:table>"
        $xmltext -replace $searchstr, $newt
    }

    function CombineExceedanceCSV( $qdList, $testType, $testWpct, $testBS, $testIOdepth, $suffix )
    {
        # Merge multiple exceedance CSVs into a single output file.
        # Column merge multiple CSV files into a single one.  Complicated by
        # the fact that the number of columns in each may vary.

        $csv = $global:details + "/ezfio_exceedance_" + $suffix + ".csv"
        if ( Test-Path -Path $csv ) {
            Remove-Item -Recurse -Force $csv | Out-Null
        }
        CSVInfoHeader > $csv
        $line1 = ""
        $line2 = ""
        foreach ($qd in $qdList) {
            $line1 = $line1 + "QD${qd} Read Exceedance,,QD${qd} Write Exceedance,,,"
            $line2 = $line2 + "rdusec,rdpct,wrusec,wrpct,,"
        }
        $line1 >> $csv;
        $line2 >> $csv;

        $files = @()
        foreach ($qd in $qdList) {
            $testname = TestName $testType $testWpct $testBS $qd $testIOdepth
            if ( Test-Path -Path "${testname}.exc.read.csv") {
                $r = [System.IO.File]::OpenText( "${testname}.exc.read.csv" )
            } else {
                $r = $null
            }
            if ( Test-Path -Path "${testname}.exc.write.csv") {
                $w = [System.IO.File]::OpenText( "${testname}.exc.write.csv" )
            } else {
                $w = $null
            }
            $files += , @( $r, $w )
        }
        do {
            $all_empty = $true
            $l = ""
            foreach ($fset in $files) {
                if (($fset[0] -eq $null) -or ($fset[0].EndOfStream)) {
                    $a = ","
                } else {
                    $a = $fset[0].ReadLine().Trim()
                    $all_empty = $false
                }
                if (($fset[1] -eq $null) -or ($fset[1].EndOfStream)) {
                    $b = ","
                } else {
                    $b = $fset[1].ReadLine().Trim()
                    $all_empty = $false
                }
                $l += "${a},${b},,"
            }
            $l >> $csv
        } while (-not $all_empty)
        foreach ($fset in $files) {
            if ($fset[0] -ne $null) {
                $fset[0].Close()
            }
            if ($fset[1] -ne $null) {
                $fset[1].Close()
            }
        }
        return $csv
    }

    function UpdateContentXMLToODS_text( $odssrc, $odsdest, $xmltext )
    {
        # Replace content.xml in an ODS file with in-memory, modified copy and
        # write new ODS. Can't just copy source.zip and replace one file, the
        # output ZIP file is not correct in many cases (opens in Excel but fails
        # ODF validation and LibreOffice fails to load under Windows)

        if (test-path $odsdest) { Remove-Item $odsdest }

        # Windows ZipArchive will not use "Store" even if we select no compression
        # so we need to have a mimetype.zip file encoded below to match ODF spec:
        $mimetypezip = @'
UEsDBBQAAAgAAICyN0+FbDmKLgAAAC4AAAAIAAAAbWltZXR5cGVhcHBsaWNhdGlvbi92bmQub2Fz
aXMub3BlbmRvY3VtZW50LnNwcmVhZHNoZWV0UEsBAhQAFAAACAAAgLI3T4VsOYouAAAALgAAAAgA
AAAAAAAAAAAAAAAAAAAAAG1pbWV0eXBlUEsFBgAAAAABAAEANgAAAFQAAAAAAA==
'@
        $bytes = [System.Convert]::FromBase64String( $mimetypezip )
        [io.file]::WriteAllBytes( $odsdest, $bytes )

        $zasrc = [System.IO.Compression.ZipFile]::Open( $odssrc, [System.IO.Compression.ZipArchiveMode]::Read )
        $zadst = [System.IO.Compression.ZipFile]::Open( $odsdest, [System.IO.Compression.ZipArchiveMode]::Update )
        foreach ($entry in $zasrc.Entries) {
            if (($entry.FullName -eq "mimetype") -or $entry.FullName.StartsWith("Thumbnails") -or $entry.FullName.StartsWith("ObjectReplacement")) {
                # Skip binary versions, and the copied-over mimetype
                continue
            }
            $newentry = $zadst.CreateEntry( $entry )
            if ($entry.FullName.EndsWith("/") -or $entry.FullName.EndsWith("\")) {
                # Directory, don't copy anything
            } elseif ($entry.FullName -eq "content.xml") {
                # Copying data for content.xml from new data
                $wr = New-Object System.IO.StreamWriter( $newentry.Open() )
                $wr.Write( $xmltext )
                $wr.Close()
            } elseif ($entry.FullName -like "Object */content.xml") {
                # Remove <table:table table:name="local-table"> table
                $rd = New-Object System.IO.StreamReader( $entry.Open() )
                $rdbytes = $rd.ReadToEnd()
                $wr = New-Object System.IO.StreamWriter( $newentry.Open() )
                $wrbytes = $rdbytes -replace "<table:table table:name=`"local-table`">.*</table:table>", ""
                $wr.write( $wrbytes )
                $wr.Close()
                $rd.Close()
            } elseif ($entry.FullName -eq "META-INF/manifest.xml") {
                # Remove ObjectReplacements from the list
                $rd = New-Object System.IO.StreamReader( $entry.Open() )
                $wr = New-Object System.IO.StreamWriter( $newentry.Open() )
                $rdbytes = $rd.ReadToEnd()
                $lines = $rdbytes.Split("`n")
                foreach ($line in $lines) {
                    if ( -not ( ($line -contains "ObjectReplacement") -or ($line -contains "Thumbnails") ) ) {
                        $wr.Write($line)
                        $wr.Write("`n")
                    }
                }
                $wr.Close()
                $rd.Close()
            } else {
                # Copying data for from the source ZIP
                $wr = New-Object System.IO.StreamWriter( $newentry.Open() )
                $rd = New-Object System.IO.StreamReader( $entry.Open() )
                $wr.Write( $rd.ReadToEnd() )
                $wr.Close()
                $rd.Close()
            }
        }
        $zadst.Dispose()
        $zasrc.Dispose()
    }

    # Use text magic and not XML editing as the XML processor doesn't seem to
    # escape special characters in the same way that OpenOffice does, leading to
    # occasional problems.  Also allows same logic to run under Linux w/sed
    [string]$xmlsrc = GetContentXMLFromODS $global:odssrc
    $xmlsrc = ReplaceSheetWithCSV_regex Timeseries $global:timeseriescsv $xmlsrc
    $xmlsrc = ReplaceSheetWithCSV_regex TimeseriesCLAT $global:timeseriesclatcsv $xmlsrc
    $xmlsrc = ReplaceSheetWithCSV_regex TimeseriesSLAT $global:timeseriesslatcsv $xmlsrc
    $xmlsrc = ReplaceSheetWithCSV_regex Tests $global:testcsv $xmlsrc
    # Potentially add exceedance data if we have it
    if ($global:fioOutputFormat -eq "json+") {
        $csv = CombineExceedanceCSV @(1, 4, 16, 32) "Rand" 30 4096 1 "exceedance30"
        $xmlsrc = ReplaceSheetWithCSV_regex Exceedance $csv $xmlsrc
    }
    # Remove draw:image references to deleted binary previews
    $xmlsrc = $xmlsrc -replace "<draw:image.*?/>",""
    $xmlsrc = $xmlsrc -replace "_DRIVE",$global:physDrive -replace "_TESTCAP",$global:testcapacity -replace "_MODEL",$global:model -replace "_SERIAL",$global:serial -replace "_OS",$global:os -replace "_FIO",$global:fioVerString
    UpdateContentXMLToODS_text $global:odssrc $global:odsdest $xmlsrc
}



function CreateIcon()
{
    $iconb64 = @'
AAABAAEAICAAAAEAIABDAgAAFgAAAIlQTkcNChoKAAAADUlIRFIAAAAgAAAAIAgGAAAAc3p6
9AAAAgpJREFUWIXtlz9rVEEUxX+zTqW9gn8Sm6ikErTQJPoB3E4UIogQ3WiiqIh/CGJhsYkG
gmCjxqwaYqO9H0DcqI1gJ2YNSLTKB9BCcmcs4oxv3s7ukpDNFHrgwbwz975z5s5w33vKWgvA
4bH3ReAisBfYQnuwCHwEym9uHnwLoKy19JVn7wAjbRJthJHZW33jquf26yLwap3FAQQ4oI2R
CwnEATYAV7UR2ZfIAECvtkY2JzSwQxtZSqgP2hhJa8CmNmAktYF/vgKrOQOfHp6q47qHZzzf
PTzTNC4LtXvwqV2pgc+PB5aTlQp492JTSgXj/Pyes888F5yBWqXkx7tKlTouyzt0nZlqaTgf
Mzd12nMFK4IV8ULOca1SworQNTAZrOLLk8FgRbVKycc2gtOIcQVjhGwVXJkA5qeHMEai5XT3
7pqfHmpoIK+R5bQxYSvOi+Tnd5683/Q+bqC+3TtOx0rXeeKeH399fikwFitlK8RiHOcPYUf/
BN9eXAu2oKN/4m/CHz7LLa8kbiD2vOxCXF7BiOCu7cfHg339/vJ6Uw4glu/4fK6b23bsrs9R
W4+OrrgPrCXSt+LkL6P/3wPJt8DI0iLt+xVrhQVtjXwAiokMVLURGQOOAKpV9BrjF1BW1lo2
HTp/AxgF9DqKX/5RffBIuV69sefcfuAK0At0tkl4AagC5Z/vJucAfgOSfC+wPSfmJAAAAABJ
RU5ErkJggg==
'@
    # Load the icon as a bitmap for user
    $iconBitmap = New-Object System.Windows.Media.Imaging.BitmapImage
    $iconBitmap.BeginInit()
    $iconBitmap.StreamSource = [System.IO.MemoryStream][System.Convert]::FromBase64String($iconb64)
    $iconBitmap.EndInit()
    $iconBitmap.Freeze()
    return $iconBitmap
}



$global:fio = ""          # FIO executable
$global:fioVerString = "" # FIO self-reported version
$global:fioOutputFormat = "json" # Can we make exceedance charts using JSON+ output?
$global:physDrive = ""    # Device path to test
$global:utilization = ""  # Device utilization % 1..100
$global:yes = $false      # Skip user verification
$global:nullio = $false   # Use the null IO engine, no real transfers done
$global:fastPrecond = $false  # Only do one sequential fill, no other preconditioning
$global:ioengine = "windowsaio"   # FIO engine to use for simplicity
$global:quickie = $false   # Do short shadown test, non-standard

$global:cpu = ""         # CPU model
$global:cpuCores = ""    # # of cores (including virtual)
$global:cpuFreqMHz = ""  # "Nominal" speed of CPU
$global:uname = ""       # Kernel name/info

$global:physDriveGiB = ""  # Disk size in GiB (2^n)
$global:physDriveGB = ""   # Disk size in GB (10^n)
$global:physDriveBase = "" # Basename (ex: nvme0n1)
$global:testcapacity = ""  # Total GiB to test
$global:model = ""         # Drive model name
$global:serial = ""        # Drive serial number

$global:ds = ""  # Datestamp to appent to files/directories to uniquify

$global:details = ""       # Test details directory
$global:testcsv = ""       # Intermediate test output CSV file
$global:timeseriescsv = "" # Intermediate iostat output CSV file
$global:timeseriesclatcsv = "" # Intermediate iostat output CSV file
$global:timeseriesslatcsv = "" # Intermediate iostat output CSV file

$global:odssrc = ""  # Original ODS spreadsheet file
$global:odsdest = "" # Generated results ODS spreadsheet file

$global:oc = New-Object System.Collections.ObjectModel.ObservableCollection[Object] # The list of tests to run

$global:iconBitmap = CreateIcon
$global:scriptName = $MyInvocation.MyCommand.Name

CheckAdmin
ParseArgs
FindFIO
CheckFIOVersion
CollectSystemInfo
CollectDriveInfo
VerifyContinue
SetupFiles

# $globals == The "global" variables to pass into the FIO runner script
$global:globals  = "`$fio = `"$global:fio`";"
$global:globals += "`$fioOutputFormat = `"$global:fioOutputFormat`";"
$global:globals += "`$physDrive = `"$global:physDrive`";"
$global:globals += "`$testcapacity = `"$global:testcapacity`";"
$global:globals += "`$timeseriescsv = `"$global:timeseriescsv`";"
$global:globals += "`$timeseriesclatcsv = `"$global:timeseriesclatcsv`";"
$global:globals += "`$timeseriesslatcsv = `"$global:timeseriesslatcsv`";"
$global:globals += "`$testcsv = `"$global:testcsv`";"
$global:globals += "`$physDriveBase = `"$global:physDriveBase`";"
$global:globals += "`$physDriveNo = `"$global:physDriveNo`";"
$global:globals += "`$details= `"$global:details`";"
$global:globals += "`$ds = `"$global:ds`";"
$global:globals += "`$ioengine = `"$global:ioengine`";"
if ( $globals:quickie ) {
    $global:globals += "`$quickie = 1;"
} else {
    $global:globals += "`$quickie = 0;"
}


DefineTests
if ($global:testmode -eq "cli") { RunAllTestsCLI }
else { RunAllTests }
# GenerateResultODS # Done in the RunAllTests function
