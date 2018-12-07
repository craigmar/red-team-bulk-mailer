param (
    [string]$infile
 )
 
#classic global variables by craig
$logfile_base = "rt-mail-send.log";


$csv = Import-Csv $infile -Header from,to,subject,message_file,server,ssl,priority,attachment,sub1,sub2,sub3,sub4,sub5 | Select-Object -Skip 1 | Foreach-Object{
    $mail_template = [IO.File]::ReadAllText($_.message_file)
	$mail_message = $mail_template.replace("<sub1>",$_.sub1).replace("<sub2>",$_.sub2).replace("<sub3>",$_.sub3).replace("<sub4>",$_.sub4).replace("<sub5>",$_.sub5)
	
	if($_.ssl -eq "y"){
	
		if($_.attachment -ne "none"){
			Send-MailMessage -To $_.to -From $_.from -Subject $_.subject -Body $mail_message -BodyAsHtml -Priority $_.priority -SmtpServer $_.server -UseSsl -Attachment $_.attachment
			$output =  (get-date).ToUniversalTime().ToString() + "," + $_.to + "," + $_.from + "," + $_.subject + "," + $_.message_file + "," + $_.server + "," + $_.attachment 
			$output >> $logfile_base
		}else{
			Send-MailMessage -To $_.to -From $_.from -Subject $_.subject -Body $mail_message -BodyAsHtml -Priority $_.priority -SmtpServer $_.server -UseSsl
			$output =  (get-date).ToUniversalTime().ToString() + "," + $_.to + "," + $_.from + "," + $_.subject + "," + $_.message_file + "," + $_.server + "," + $_.attachment 
			$output >> $logfile_base
		}
	}else{
	
		if($_.attachment -ne "none"){
			Send-MailMessage -To $_.to -From $_.from -Subject $_.subject -Body $mail_message -BodyAsHtml -Priority $_.priority -SmtpServer $_.server -Attachment $_.attachment
			$output =  (get-date).ToUniversalTime().ToString() + "," + $_.to + "," + $_.from + "," + $_.subject + "," + $_.message_file + "," + $_.server + "," + $_.attachment 
			$output >> $logfile_base
		}else{
			Send-MailMessage -To $_.to -From $_.from -Subject $_.subject -Body $mail_message -BodyAsHtml -Priority $_.priority -SmtpServer $_.server
			$output =  (get-date).ToUniversalTime().ToString() + "," + $_.to + "," + $_.from + "," + $_.subject + "," + $_.message_file + "," + $_.server + "," + $_.attachment 
			$output >> $logfile_base
		}
	
	}	
}
