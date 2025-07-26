# script-printerinstall.ps1
# ==================================================
# Description
# ==================================================
# Usage
# ==================================================


Write-Host "Starting script-printerinstall.ps1 ..."

# include
# . ".\functions.ps1"
# . ".\urls.ps1"
# . ".\values.ps1"

# var
# $var = ""

# Documentation
# options on Printer Port page
# 
# Port Name
# $printerport_port_name = ""
# 
# Printer Name or IP Address
# $printerport_printer_name_or_ip_address = ""
# 
# Protocol
# $printerport_protocol = "Raw"
# takes values
# "Raw"
# "LPR"
# 
# switch ($protocol) {
    # "Raw" {
        # # "Raw"
        # 
        # Raw Settings
        # 
        # Port Number:
        # $printerport_raw_port_number = "9100"
        # 
        # break
    # }
    # "LPR" {
        # # "LPR"
        # 
        # LPR Settings
        # 
        # Queue Name:
        # $printerport_lpr_queue_name = "BINARY_P1"
        # 
        # LPR Byte Counting Enabled
        # $printerport_lpr_byte_counting_enabled = $false
        # takes values
        # $true
        # $false
        # 
        # break
    # }
    # default {
        # # Default
        # Write-Host "error"
        # break
    # }
# }
# 
# SNMP Status Enabled
# $printerport_snmp_status_enabled = $true
# takes values
# $true
# $false
# 
# if ($printerport_snmp_status_enabled) {
    # 
    # Community Name:
    # $printerport_snmp_community_name = "public"
    # 
    # SNMP Device Index:
    # $printerport_snmp_device_index = "1"
    # 
# }
# 
# 
# Common Settings
# $printerport_raw_port_number = "9100"
# $printerport_lpr_queue_name = "BINARY_P1"
# $printerport_snmp_community_name = "public"
# $printerport_snmp_device_index = "1"
# 
# 
# powershell commands corresponding to fields
# 
# Add-PrinterPort
# -Name $printerport_port_name
# special_note: $printerport_printer_name_or_ip_address is assigned to -PrinterHostAddress or -LprHostAddress depending on Raw/LPR
# if (raw) {
    # -PrinterHostAddress $printerport_printer_name_or_ip_address
    # -PortNumber $printerport_raw_port_number
# } else { # lpr
    # -LprHostAddress $printerport_printer_name_or_ip_address
    # -LprQueueName $printerport_lpr_queue_name
    # -LprByteCounting
# }
# if (snmp) {
    # -SNMP $printerport_snmp_device_index
    # -SNMPCommunity $printerport_snmp_community_name
# }

# Cases
Function ListCases() {
    # list 6 cases
    Write-Host "List of 6 Cases of Printer Ports"
    Write-Host "Case 1: Raw, SNMP false"
    Write-Host "Case 2: Raw, SNMP true"
    Write-Host "Case 3: LPR, ByteCounting false, SNMP false"
    Write-Host "Case 4: LPR, ByteCounting false, SNMP true"
    Write-Host "Case 5: LPR, ByteCounting true, SNMP false"
    Write-Host "Case 6: LPR, ByteCounting true, SNMP true"
}

Function DetermineCase($printerport_protocol, $printerport_lpr_byte_counting_enabled, $printerport_snmp_status_enabled) {
    # check Raw or LPR
    if ($printerport_protocol -eq "Raw") {
        # check SNMP
        if (-Not($printerport_snmp_status_enabled)) {
            $printerport_case_number = 1
        } elseif ($printerport_snmp_status_enabled) {
            $printerport_case_number = 2
        }
    } elseif ($printerport_protocol -eq "LPR") {
        # check LPR Byte Counting
        if (-Not($printerport_lpr_byte_counting_enabled)) {
            # check SNMP
            if (-Not($printerport_snmp_status_enabled)) {
                $printerport_case_number = 3
            } elseif ($printerport_snmp_status_enabled) {
                $printerport_case_number = 4
            }
        } elseif ($printerport_lpr_byte_counting_enabled) {
            # check SNMP
            if (-Not($printerport_snmp_status_enabled)) {
                $printerport_case_number = 5
            } elseif ($printerport_snmp_status_enabled) {
                $printerport_case_number = 6
            }
        }
    }
    # return case number
    return $printerport_case_number
}

# get path of CSV file
Write-Host "CSV file - list of printers"
if (Test-Path -Path "$HOME\Downloads\printerslist.csv") {
    # use this
    $path_csv_listofprinters = "$HOME\Downloads\printerslist.csv"
} else {
    # ask
    Add-Type -AssemblyName "System.Windows.Forms"
    $fileBrowser = New-Object -TypeName "System.Windows.Forms.OpenFileDialog" -Property @{ InitialDirectory = "$HOME\Downloads" }
    $null = $fileBrowser.ShowDialog()
    $path_file = $fileBrowser.FileName
    if ($path_file -eq "") {
        $path_csv_listofprinters = "$HOME\Downloads\printerslist.csv"
    } else {
        $path_csv_listofprinters = $path_file
    }
    Clear-Variable path_file
}

# get path of drivers folder
Write-Host "Drivers folder of printers"
if (Test-Path -Path "$HOME\Downloads\PrinterDrivers") {
    # use this
    $path_folder_driversinf = "$HOME\Downloads\PrinterDrivers"
} else {
    # ask
    Add-Type -AssemblyName "System.Windows.Forms"
    $folderBrowser = New-Object -TypeName "System.Windows.Forms.FolderBrowserDialog"
    $null = $folderBrowser.ShowDialog()
    $path_folder = $folderBrowser.SelectedPath
    if ($path_folder -eq "") {
        $path_folder_driversinf = "$HOME\Downloads\PrinterDrivers"
    } else {
        $path_folder_driversinf = $path_folder
    }
    Clear-Variable path_folder
}

# import data from CSV
# 
# columns in CSV
# printer_name_official
# printerport_printer_name_or_ip_address
# printerport_protocol
# printerport_raw_port_number
# printerport_lpr_queue_name
# printerport_lpr_byte_counting_enabled
# printerport_snmp_status_enabled
# printerport_snmp_community_name
# printerport_snmp_device_index
# printer_drivername_in_inf
# org_name
# 
$importedcsv = Import-Csv -Path $path_csv_listofprinters
$number_of_printers_in_csv = $importedcsv.count
# list
for ( $i = 0 ; $i -lt $number_of_printers_in_csv ; $i++ ) {
    Write-Host "$($i+1) : $($importedcsv[$i].org_name) - $($importedcsv[$i].printer_name_official)"
}
# give one more option for manual input
Write-Host "$($number_of_printers_in_csv+1) : Manual input"
# Ask
Do {
    [int]$selected_printer_number = Read-Host -Prompt "Enter number - Known printer (Leave blank if printer not on list): "
} while ( ($selected_printer_number -lt 1) -or ($selected_printer_number -gt ($number_of_printers_in_csv+1)) )

# Populate variables
if ($selected_printer_number -ne ($number_of_printers_in_csv+1)) {
    $printer_name_official = $importedcsv[$selected_printer_number-1].printer_name_official
    $printerport_printer_name_or_ip_address = $importedcsv[$selected_printer_number-1].printerport_printer_name_or_ip_address
    $printerport_protocol = $importedcsv[$selected_printer_number-1].printerport_protocol
    [int]$printerport_raw_port_number = $importedcsv[$selected_printer_number-1].printerport_raw_port_number
    $printerport_lpr_queue_name = $importedcsv[$selected_printer_number-1].printerport_lpr_queue_name
    [bool]$printerport_lpr_byte_counting_enabled = [System.Convert]::ToBoolean($importedcsv[$selected_printer_number-1].printerport_lpr_byte_counting_enabled)
    [bool]$printerport_snmp_status_enabled = [System.Convert]::ToBoolean($importedcsv[$selected_printer_number-1].printerport_snmp_status_enabled)
    $printerport_snmp_community_name = $importedcsv[$selected_printer_number-1].printerport_snmp_community_name
    [int]$printerport_snmp_device_index = $importedcsv[$selected_printer_number-1].printerport_snmp_device_index
    $printer_drivername_in_inf = $importedcsv[$selected_printer_number-1].printer_drivername_in_inf
    $org_name = $importedcsv[$selected_printer_number-1].org_name
} else {
    # manual input
    $printer_name_official = Read-Host -Prompt "Enter printer_name_official"
    $printerport_printer_name_or_ip_address = Read-Host -Prompt "Enter printerport_printer_name_or_ip_address"
    $printerport_protocol = Read-Host -Prompt "Enter printerport_protocol"
    [int]$printerport_raw_port_number = Read-Host -Prompt "Enter printerport_raw_port_number"
    $printerport_lpr_queue_name = Read-Host -Prompt "Enter printerport_lpr_queue_name"
    $printerport_lpr_byte_counting_enabled = Read-Host -Prompt "Enter printerport_lpr_byte_counting_enabled"
    [System.Convert]::ToBoolean($printerport_lpr_byte_counting_enabled)
    $printerport_snmp_status_enabled = Read-Host -Prompt "Enter printerport_snmp_status_enabled"
    [System.Convert]::ToBoolean($printerport_snmp_status_enabled)
    $printerport_snmp_community_name = Read-Host -Prompt "Enter printerport_snmp_community_name"
    [int]$printerport_snmp_device_index = Read-Host -Prompt "Enter printerport_snmp_device_index"
    $printer_drivername_in_inf = Read-Host -Prompt "Enter printer_drivername_in_inf"
    $org_name = Read-Host -Prompt "Enter org_name"
}

# calculate derived values
$printerport_port_name = $printerport_printer_name_or_ip_address + "_$org_name"
$printer_displayname_final = $org_name + " " + $printer_name_official
[int]$printerport_case_number = DetermineCase $printerport_protocol $printerport_lpr_byte_counting_enabled $printerport_snmp_status_enabled
if ($null -eq $printerport_case_number) {

    ListCases
    Write-Host " "

    Do {
        [int]$printerport_case_number = Read-Host -Prompt "Enter number - Case of printer port: "
    } while ( ($printerport_case_number -lt 1) -or ($printerport_case_number -gt 6) )
}

Write-Host " "
Write-Host " "

Write-Host "PrinterPort Case Number: $printerport_case_number"

Write-Host " "
Write-Host " "

# Install

# add inf files into DriverStore
# ask if added
Do {
    [char]$drivers_added_into_driverstore = Read-Host -Prompt "Enter letter - Are drivers added into driver store already? y , n "
} while ( ($drivers_added_into_driverstore -ne "y") -and ($drivers_added_into_driverstore -ne "n") )
# add if no
if ($drivers_added_into_driverstore -eq "n") {
    # check if path exist
    if (Test-Path -Path $path_folder_driversinf) {
        Write-Host "Adding inf files into DriverStore ..."
        Get-ChildItem -Path $path_folder_driversinf -Recurse -Filter "*.inf" | ForEach-Object { pnputil.exe /add-driver $_.FullName }
        Write-Host "... Done"
    } else {
        Write-Host "No driver folder found"
        Write-Host "Place drivers in HOME\Downloads\PrinterDrivers"
        exit
    }
}
else {
    Write-Host "Inf files already added into driver store, proceeding to next step ... "
}

Write-Host " "

# install driver from DriverStore
Write-Host "Installing driver from DriverStore ..."
Add-PrinterDriver -Name $printer_drivername_in_inf
Write-Host "... Done"

Write-Host " "

# Add-PrinterPort to add printer port
Write-Host "Adding printer port ..."
switch ($printerport_case_number)
{
    1 {
        # 1
        Add-PrinterPort -Name $printerport_port_name -PrinterHostAddress $printerport_printer_name_or_ip_address -PortNumber $printerport_raw_port_number
        break
    }
    2 {
        # 2
        Add-PrinterPort -Name $printerport_port_name -PrinterHostAddress $printerport_printer_name_or_ip_address -PortNumber $printerport_raw_port_number -SNMP $printerport_snmp_device_index -SNMPCommunity $printerport_snmp_community_name
        break
    }
    3 {
        # 3
        Add-PrinterPort -Name $printerport_port_name -LprHostAddress $printerport_printer_name_or_ip_address -LprQueueName $printerport_lpr_queue_name
        break
    }
    4 {
        # 4
        Add-PrinterPort -Name $printerport_port_name -LprHostAddress $printerport_printer_name_or_ip_address -LprQueueName $printerport_lpr_queue_name -SNMP $printerport_snmp_device_index -SNMPCommunity $printerport_snmp_community_name
        break
    }
    5 {
        # 5
        Add-PrinterPort -Name $printerport_port_name -LprHostAddress $printerport_printer_name_or_ip_address -LprQueueName $printerport_lpr_queue_name -LprByteCounting
        break
    }
    6 {
        # 6
        Add-PrinterPort -Name $printerport_port_name -LprHostAddress $printerport_printer_name_or_ip_address -LprQueueName $printerport_lpr_queue_name -LprByteCounting -SNMP $printerport_snmp_device_index -SNMPCommunity $printerport_snmp_community_name
        break
    }
    default {
        # default
        Write-Host "Error"
        break
    }
}
Write-Host "... Done"

Write-Host " "

# Add-Printer into Printers list
Write-Host "Adding printer ..."
Add-Printer -Name $printer_displayname_final -DriverName $printer_drivername_in_inf -PortName $printerport_port_name
Write-Host "... Done"

Write-Host " "

# Show Devices and Printers
Show-ControlPanelItem -CanonicalName "Microsoft.DevicesAndPrinters"


Write-Host ""

Write-Host "Terminating script-printerinstall.ps1 ..."
# pause


# ==================================================
# Notes
# ==================================================
