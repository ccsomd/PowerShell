# EMPLOYEE VARIABLES
# ========================================================================
# ========================================================================
$year                           = "2022"
$employee                       = ""
$date_hired                     = ""
$carry_over_leave_annual        = ""
$carry_over_leave_sick          = ""
$carry_over_leave_comp          = ""
$carry_over_leave_holiday       = ""
$carry_over_leave_new_holiday   = ""
$carry_over_leave_personal      = ""


# LEAVE RECORD TEMPLATES
# ========================================================================
# ========================================================================
$template_location      = "C:\Code\Leave Records\Templates\2022\"

$template_corrections   = "2022 CORRECTIONS MASTER.xls"
$template_critical      = "2022 CRITICAL MASTER.xls"
$template_ftrh          = "2022 FT-REDUCED HOURS MASTER.xls"
$template_non_critical  = "2022 NON-CRITICAL MASTER.xls"
$template_operational   = "2022 OPERATIONAL MASTER.xls"
$template_part_time     = "2022 PART-TIME MASTER.xls"
$template_pco           = "2022 PCO MASTER.xls"
$template_sworn         = "2022 SWORN MASTER.xls"


# LEAVE RECORD FOLDERS
# ========================================================================
# ========================================================================
$leave_record_source_folder        = "C:\Code\Leave Records\2021"
$leave_record_destination_folder   = "C:\Code\Leave Records\2022"

Write-Host ""
Write-Host "Leave Records Location: $leave_record_source_folder"
Write-Host "Leave Records Destination: $leave_record_destination_folder"
Write-Host ""

# CHANGE SYSTEM DATE TO CURRENT YEAR
# =============================================
# =============================================
$system_date = "01/01/2022 11:00"


# DUPLICATE LEAVE RECORD FOLDER STRUCTURE
# ========================================================================
# ========================================================================

# Remove existing destination folder
Remove-Item $leave_record_destination_folder -Recurse

# Create new destination folder
New-Item -ItemType Directory -Path $leave_record_destination_folder

# Write-Host $leave_record_destination_folder

# Copy source folder subfolders to destination folder
Get-ChildItem $leave_record_source_folder -Attributes D -Recurse | ForEach-Object {

    # Copy Level 1 Subfolders
    if($_.Parent.FullName -eq $leave_record_source_folder) {
        New-Item -ItemType Directory -Path ($leave_record_destination_folder + "\$_")
    }

    # Copy Level 2 Subfolders
    if($_.Parent.Parent.FullName -eq $leave_record_source_folder) {
        New-Item -ItemType Directory -Path ($leave_record_destination_folder + "\" + $_.Parent.Name + "\$_")
    }

    # Copy Level 3 Subfolders
    if($_.Parent.Parent.Parent.FullName -eq $leave_record_source_folder) {
        New-Item -ItemType Directory -Path ($leave_record_destination_folder + "\" + $_.Parent.Parent.Name + "\" + $_.Parent.Name + "\$_")
    }
}


# CHANGE SYSTEM TIME TO JANUARY OF CURRENT YEAR
# =============================================
# =============================================
set-date -date $system_date | Out-Null


# EXCEL OBJECT
# =============================
# =============================
$excel = New-Object -ComObject excel.application
$excel.Visible = $false
$excel.DisplayAlerts = $false



# LOOP OVER ALL OF THE EXCEL SPREADSHEETS IN THE SOURCE FOLDER
# ========================================================================
# ========================================================================
Get-ChildItem $leave_record_source_folder -File -Recurse -Filter *.xls |

ForEach-Object {

    # CHANGE SYSTEM TIME TO JANUARY
    # =============================
    # =============================
    set-date -date $system_date | Out-Null

    # Zero Out Variables
    # ==================
    $carry_over_leave_annual        = ""
    $carry_over_leave_annual_excess = ""
    $carry_over_leave_sick          = ""
    $carry_over_leave_comp          = ""
    $carry_over_leave_holiday       = ""
    # $carry_over_leave_new_holiday   = ""
    $carry_over_leave_personal      = ""

    # CALCULATE THE DESTINATION BASED ON THE LOCATION
    # ========================================================================
    # ========================================================================
    $location = $_.FullName
    $destination = $location.Replace("$leave_record_source_folder","$leave_record_destination_folder")




# ========================================================================
# ========================================================================
# We have the location of the file
# We have the destination of the file
# We know what template to use
# Let's begin
# ========================================================================
# ========================================================================


    # OPEN EMPLOYEE LEAVE RECORD
    # ========================================================================
    # ========================================================================
    $wb = $excel.Workbooks.Open($_.FullName)
    $ws = $wb.Sheets.Item(1)
    $wt = $wb.Sheets.Item(3)


    # PREPARE TO GET EMPLOYEE TEMPLATE CODE - IF CODE IS NOT LOCATED AT 40,3
    # ========================================================================
    # ========================================================================
    if ($ws.Cells.Item(40,3).text -eq "") {
        if ($ws.Cells.Item(34,3).text -ne "") {
            $ws.Cells.Item(40,3) = $ws.Cells.Item(34,3).text
        } elseif ($ws.Cells.Item(35,3).text -ne "") {
            $ws.Cells.Item(40,3) = $ws.Cells.Item(35,3).text
        } elseif ($ws.Cells.Item(39,3).text -ne "") {
            $ws.Cells.Item(40,3) = $ws.Cells.Item(39,3).text
        } else {
            Write-Host "!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!"
            Write-Host "Template Code Error"
            Write-Host $_.FullName
            BREAK
        }
    }


    # EMPLOYEE TEMPLATE CODE
    # ========================================================================
    # ========================================================================
    $employee_template_code     = $ws.Cells.Item(40,3).text


    # EMPLOYEE INFO
    # ========================================================================
    # ========================================================================
    $employee                   = $ws.Cells.Item(5,3).text
    $date_hired                 = $ws.Cells.Item(6,3).text

    switch ($employee_template_code) {
        "Sworn" {
            $carry_over_leave_annual      = $ws.Cells.Item(18,3).text
            $carry_over_leave_sick        = $ws.Cells.Item(19,3).text
            $carry_over_leave_comp        = $ws.Cells.Item(20,3).text
            $carry_over_leave_holiday     = $ws.Cells.Item(21,3).text
            # $carry_over_leave_new_holiday = $wt.Cells.Item(15,22).text
            $carry_over_leave_personal    = $wt.Cells.Item(15,19).text
        }
        "Corrections" {
            $carry_over_leave_annual      = $ws.Cells.Item(17,3).text
            $carry_over_leave_sick        = $ws.Cells.Item(18,3).text
            $carry_over_leave_comp        = $ws.Cells.Item(19,3).text
            $carry_over_leave_holiday     = $ws.Cells.Item(20,3).text
            # $carry_over_leave_new_holiday = $wt.Cells.Item(15,19).text
        }
        default {
            $carry_over_leave_annual    = $ws.Cells.Item(17,3).text
            $carry_over_leave_sick      = $ws.Cells.Item(18,3).text
            $carry_over_leave_comp      = $ws.Cells.Item(19,3).text
            $carry_over_leave_holiday   = $ws.Cells.Item(20,3).text
        }
    }

    # CAP HOLIDAY HOURS (FOR 2019 -> 2020)
    # ========================================================================
    # ========================================================================
    # if ($employee_template_code -eq "Sworn" -or $employee_template_code -eq "Corrections") {
    #     if (($carry_over_leave_holiday/1) -gt 15) {
    #         $carry_over_leave_holiday = 15
    #     }
    #     if (($carry_over_leave_new_holiday/1) -gt 4) {
    #         $carry_over_leave_new_holiday = 4
    #     }

    #     $carry_over_leave_holiday = ($carry_over_leave_holiday/1) + ($carry_over_leave_new_holiday/1)

    #     if ($employee_template_code -eq "Sworn") {
    #         $carry_over_leave_holiday = ($carry_over_leave_holiday/1) * 10
    #     } else {
    #         $carry_over_leave_holiday = ($carry_over_leave_holiday/1) * 8.5
    #     }
    # }

    # CAP HOLIDAY HOURS - SWORN
    # ========================================================================
    # ========================================================================
    if ($employee_template_code -eq "Sworn" -and ($carry_over_leave_holiday/1) -gt 230) {
        $carry_over_leave_holiday = 230
    }

    # CAP HOLIDAY HOURS - CORRECTIONS
    # ========================================================================
    # ========================================================================
    if ($employee_template_code -eq "Corrections" -and ($carry_over_leave_holiday/1) -gt 178.5) {
        $carry_over_leave_holiday = 178.5
    }


    # CALCULATE CARRY OVER ANNUAL LEAVE
    # ========================================================================
    # ========================================================================
    # Write-Host ""
    # Write-Host "Employee: $employee"
    # Write-Host "Annual Start: $carry_over_leave_annual"
    # Write-Host "Sick Start: $carry_over_leave_sick"

    if (($employee_template_code -eq "Operational") -or ($employee_template_code -eq "Non-Critical") -or ($employee_template_code -eq "FTRH")) {
        if (($carry_over_leave_annual/1) -gt 360) {
            $carry_over_leave_annual_excess = (($carry_over_leave_annual/1) - 360)
            $carry_over_leave_sick = (($carry_over_leave_sick/1) + $carry_over_leave_annual_excess)
            $carry_over_leave_annual = 360
        }
    } else {
        if (($carry_over_leave_annual/1) -gt 720) {
            $carry_over_leave_annual_excess = (($carry_over_leave_annual/1) - 720)
            $carry_over_leave_sick = (($carry_over_leave_sick/1) + $carry_over_leave_annual_excess)
            $carry_over_leave_annual = 720
        }
    }

    # Write-Host "Annual After: $carry_over_leave_annual"
    # Write-Host "Annual Excess: $carry_over_leave_annual_excess"
    # Write-Host "Sick End $carry_over_leave_sick"


    # CHOOSE THE APPROPRIATE LEAVE RECORD TEMPLATE
    # ========================================================================
    # ========================================================================
    switch ($employee_template_code) {
        "Corrections"   { $leave_template = $template_location + $template_corrections }
        "Critical"      { $leave_template = $template_location + $template_critical }
        "FTRH"          { $leave_template = $template_location + $template_ftrh }
        "Non-Critical"  { $leave_template = $template_location + $template_non_critical }
        "Operational"   { $leave_template = $template_location + $template_operational }
        "Part-Time"     { $leave_template = $template_location + $template_part_time }
        "PCO"           { $leave_template = $template_location + $template_pco }
        "Sworn"         { $leave_template = $template_location + $template_sworn }
        Default         { $leave_template = "" }
    }


    $excel.Workbooks.Close()


    # Write-Host ""
    Write-Host "Employee: $employee"
    # Write-Host "Date Hired: $date_hired"
    # Write-Host "Annual: $carry_over_leave_annual"
    # Write-Host "Sick: $carry_over_leave_sick"
    # Write-Host "Comp: $carry_over_leave_comp"
    # Write-Host "Holiday: $carry_over_leave_holiday"
    # Write-Host "Personal: $carry_over_leave_personal"
    # Write-Host ""
    # Write-Host "Template: $employee_template_code"
    # Write-Host "Template File: $leave_template"


    # OPEN CURRENT LEAVE RECORD TEMPLATE
    # ========================================================================
    # ========================================================================
    $wb = $excel.Workbooks.Open($leave_template)
    $ws = $wb.Sheets.Item(1)

    # UPDATE CELLS
    # ========================================================================
    # ========================================================================
    $ws.Cells.Item(4,3) = $year
    $ws.Cells.Item(5,3) = $employee
    $ws.Cells.Item(6,3) = $date_hired
    $ws.Cells.Item(10,3) = $carry_over_leave_annual
    $ws.Cells.Item(11,3) = $carry_over_leave_sick
    $ws.Cells.Item(12,3) = $carry_over_leave_comp
    $ws.Cells.Item(13,3) = $carry_over_leave_holiday

    if ($employee_template_code -eq "Sworn") {
        $ws.Cells.Item(14,3) = $carry_over_leave_personal
    }

    # SAVE SPREADSHEET
    # ================
    $ws = $wb.Sheets.Item(4)
    $ws.activate()
    $excel.ActiveWorkbook.SaveAs($destination)

    # CLOSE
    # =====
    $excel.Workbooks.Close()
}

$excel.Quit()

# W32tm /resync /force
