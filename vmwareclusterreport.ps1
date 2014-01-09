<#
	NAME: vmwareclusterreport.ps1
	AUTHOR: Michael McDowell
	DATE: 11/26/2013
	DESCRIPTION: Pulls historical metrics from the vCenter database for CPU, memory and network usage,
	creates Excel graphs and emails them to recipients configured.
#>

add-pssnapin VMware.VimAutomation.Core
#Report Settings
$cluster = 'CLUSTERNAME' # The full name of the cluster as it appears in vCenter
$start = (Get-Date).AddDays(-7) # This can be changed as required
$stats = @("cpu.usage.average", "mem.usage.average", "net.usage.average") # Additional elements can be added to this array if other stats need to be collected
$viserver = 'VCENTERSERVER' # This is the vCenter server

# Mail Settings
$sender = 'FROMEMAIL' # The FROM email address
$recievers = @('RECIEVER1','RECIEVER2','RECIEVER3') # The TO email addresses
$workdir = 'C:\SCRIPTS' # This is the directory the reports are saved to
$mailserver = 'MAILSRV' # This is the SMTP server
$subject = "Weekly VMware Report"


#########
$attachments = @()
Connect-ViServer $viserver
#Pull hosts from cluster
$hosts = Get-VMHost -Location $cluster

#Create excel COM object
$excel = New-Object -ComObject excel.application
$excel.Visible = $False
$excel.DisplayAlerts = $False

#Define function to create worksheets and charts
function get-report($s){
    # Query the data from vCenter according to the stats given
    $data = Get-Stat -Entity $h -start $start -MaxSamples 10000 -Stat $s -IntervalMins 30 | sort Timestamp
    $workbook.worksheets.add()
    $vmhostsheet = $workbook.worksheets.item(1)
    $vmhostsheet.Activate() | Out-null
    $row = 1
    $column = 1
    $initialrow = $row
    $vmhostsheet.name = $s
    $vmhostsheet.Cells.Item($row,$column) = 'Timestamp'
    $column++
    $vmhostsheet.Cells.Item($row,$column) = $s
    #Down to the next row and back to column 1
    $row++
    $column = 1
    ForEach($vmrow in $data){
        $vmhostsheet.Cells.Item($row,$column) = $vmrow.Timestamp
        $vmhostsheet.Cells.Item($row,$column).NumberFormat = 'm/d/yy h:mm;@'
        $Column++
        #Format this if it is a % based metric
        if($vmrow.Unit -eq '%'){
            $vmhostsheet.Cells.Item($row,$column) = ($vmrow.Value * .01)
            $vmhostsheet.Cells.Item($row,$column).NumberFormat = '##.##%'
        }
        else{
            $vmhostsheet.Cells.Item($row,$column) = $vmrow.Value
        }
        $Column = 1
        $row++
    }
    # Create the chart
    $chartType = [microsoft.office.interop.excel.xlChartType]::xlLine
    $range = $vmhostsheet.UsedRange
    $range.EntireColumn.AutoFit()
    $workbook.charts.add()
    $workbook.ActiveChart.chartType = $chartType
    $workbook.ActiveChart.SetSourceData($range)
    $workbook.ActiveChart.Axes(1).CategoryType = 2
    $workbook.ActiveChart.name = $s + 'graph'
}
# Run reports below 
ForEach($h in $hosts){
    # Add a workbook
    $workbook = $excel.Workbooks.Add()
    # Remove other worksheets
    1..2 | ForEach {
    $Workbook.worksheets.item(2).Delete()
    }
    ForEach($stat in $stats){
        get-report($stat)
    }
    $wbpath = $workdir + (Get-Date -Format yyyyMMdd) + ($h.Name.split('.')[0]) + '.xlsx'
    $attachments += $wbpath
    $workbook.SaveAs($wbpath)
    $workbook.close()
}
# Email attachments 
$msg = New-Object Net.Mail.MailMessage
$smtp = New-Object Net.Mail.SmtpClient($mailserver)
$msg.From = $sender
ForEach($r in $recievers){
    $msg.To.Add($r)
}
$msg.Subject = $subject
ForEach($a in $attachments){
    $msg.Attachments.Add($a)
}
$smtp.Send($msg)