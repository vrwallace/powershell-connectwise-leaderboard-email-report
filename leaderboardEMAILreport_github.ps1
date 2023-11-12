####################################
# Program: leaderboardEMAILreport.ps1
# By: Von Wallace
# To run add the following to the login script
# powershell.exe �Noninteractive �Noprofile �Command "C:\support\leaderboardEMAILreport.ps1"
# Program queries CW SQL server
###################################

[Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12

try {

    
   
    #settings
    $Version = "1.01ps"
    $smtpserver = "smtp.office365.com"
    $smtpport = "587"
    
    
    $smtpfrom = "leaderboard@somewhere.net"
    
    $smtpto = "n@somwhere.net"
    
    
    $sendusername = "leaderboard@somewhere.net"
    $sendpassword = "hardtoguesspassword"
    
    $uid = "dbuserSQLAlerts"
    $pwd = "dbpassword"
       
    $sqlserver = "ipaddress"
    $SQLDBName = "cwwebapp"

    $Connection = New-Object System.Data.SQLClient.SQLConnection
    $Connection.ConnectionString = "Server = $SQLServer; Database = $SQLDBName; User ID = $uid; Password = $pwd;"
    #$Connection.ConnectionString = "Server = $SQLServer; Database = $SQLDBName; Integrated Security = True;"
    
    $Connection.Open()
    try {
        $Command = New-Object System.Data.SQLClient.SQLCommand
        $Command.Connection = $Connection


        
    
        
        $sql = "begin`r`n"

        $report = @"
        <!DOCTYPE html>
<html lang="en">
  <head>
    <meta name=`"viewport`" content=`"width=device-width,initial-scale=1`">
    <meta charset=`"utf-8`">
    <meta name=`"generator`" content=`"CoffeeCup HTML Editor (www.coffeecup.com)`">
    <meta name=`"dcterms.created`" content=`"Wed, 05 Oct 2022 23:08:13 GMT`">
    <meta name=`"description`" content=`"`">
    <meta name=`"keywords`" content=`"`">
    <title>Leaderboard Report</title>
<style>
TABLE { font-family: Arial, Helvetica, sans-serif; text-align: center;font-weight: 400;border-width: 1px; border-style: solid; border-color: black; border-collapse: collapse;}
TH {border-width: 1px; padding: 3px; border-style: solid; border-color: black; color: white; background-color: #942925;}
TD {border-width: 1px; padding: 3px; border-style: solid; border-color: black;}
@media all and (min-width: 560px) {
    .container {
      border-radius: 8px;
      -webkit-border-radius: 8px;
      -moz-border-radius: 8px;
      -khtml-border-radius: 8px;
    }
  }
  
</style>
</head>
<body>
"@    
        $report = $report + "<center>"
        $report = $report + "<br/><br/><br/>"

        $report = $report + "<h1>Leaderboard Report</h1>"
        $report = $report + "<h4>Version: " + $version + "</h4></center><br>"

       
        $report = $report + "<h4>Leaderboard 'Closed Tickets' for this Week as of Today</h4>"      
               
        $sql = $sql + "`r`nSELECT top 100 resource, count(resource) as 'Closed Tickets',sum(Hours_bill) as 'Hours Billed' 
        FROM [cwwebapp_nct].[dbo].[v_company_time] where (( sr_status='Closed') and ( date>=DATEADD(week, DATEDIFF(week, 0, getdate()), 0) )) group by resource order by count(resource) desc;
        "


        $sql = $sql + "END`r`n"
        #write-host $sql
        # write-host $sql
        $Command.CommandText = $sql
        $datareader = $Command.ExecuteReader()
        IF ($DATAREADER.HASROWS) {
            #write-host "Has rows"



            $report = $report + "<table style=`"width:100%`" class=`"container`">
  <tr>
    
  
  

  <th style=`"width:20%`">&#9734; Ranking &#9734;</th> 
    <th>Resource</th>
    <th style=`"width:20%`">Closed Tickets</th> 
    <th style=`"width:20%`">Hours Billed</th>
    
    
  </tr>"

 

            $place = 0
            WHILE ($DATAREADER.READ()) {
                $place = $place + 1          
                
 
                $report = $report + "<tr><td>" + $place + "</td><td>" + $DATAREADER[0] + "</td><td>" + $DATAREADER[1] + "</td><td>" + $DATAREADER[2] + "</td></tr>" 
               
            }

            $report = $report + "</table>"
        }
        $datareader.close()
        $sql = "begin`r`n"
        #month start

        $report = $report + "<h4>Leaderboard 'Closed Tickets' for this Month as of Today</h4>"      
               
        $sql = $sql + "`r`nSELECT top 100 resource, count(resource) as 'Closed Tickets',sum(Hours_bill) as 'Hours Billed' 
FROM [cwwebapp_nct].[dbo].[v_company_time] where (( sr_status='Closed') and ( date>=DATEADD(month, DATEDIFF(month, 0, getdate()), 0) )) group by resource order by count(resource) desc;"



        $sql = $sql + "END`r`n"
        #write-host $sql
        # write-host $sql
        $Command.CommandText = $sql
        $datareader = $Command.ExecuteReader()
        IF ($DATAREADER.HASROWS) {
            #write-host "Has rows"



            $report = $report + "<table style=`"width:100%`" class=`"container`">
<tr>




<th style=`"width:20%`">&#9734; Ranking &#9734;</th>
<th>Resource</th>
<th style=`"width:20%`">Closed Tickets</th>  
<th style=`"width:20%`">Hours Billed</th>


</tr>"



            $place = 0
            WHILE ($DATAREADER.READ()) {
        
                $place = $place + 1

                $report = $report + "<tr><td>" + $place + "</td><td>" + $DATAREADER[0] + "</td><td>" + $DATAREADER[1] + "</td><td>" + $DATAREADER[2] + "</td></tr>" 
       
            }

            $report = $report + "</table>"
        }
        $datareader.close()
        #month end

        #year start
        $sql = "begin`r`n"
        $report = $report + "<h4>Leaderboard 'Closed Tickets' for this Year as of Today</h4>"      
               
        $sql = $sql + "`r`nSELECT top 100 resource, count(resource) as 'Closed Tickets',sum(Hours_bill) as 'Hours Billed' 
FROM [cwwebapp_nct].[dbo].[v_company_time] where (( sr_status='Closed') and ( date>=DATEADD(year, DATEDIFF(year, 0, getdate()), 0) )) group by resource order by count(resource) desc;"



        $sql = $sql + "END`r`n"
        #write-host $sql
        # write-host $sql
        $Command.CommandText = $sql
        $datareader = $Command.ExecuteReader()
        IF ($DATAREADER.HASROWS) {
            #write-host "Has rows"



            $report = $report + "<table style=`"width:100%`" class=`"container`">
<tr>




<th style=`"width:20%`">&#9734; Ranking &#9734;</th>
<th>Resource</th>
<th style=`"width:20%`">Closed Tickets</th>  
<th style=`"width:20%`">Hours Billed</th>


</tr>"



            $place = 0
            WHILE ($DATAREADER.READ()) {
        
                $place = $place + 1

                $report = $report + "<tr><td>" + $place + "</td><td>" + $DATAREADER[0] + "</td><td>" + $DATAREADER[1] + "</td><td>" + $DATAREADER[2] + "</td></tr>" 
       
            }

            $report = $report + "</table>"
        }
        $datareader.close()
        #year end




        #BILLED HOURS

        $sql = "begin`r`n"

        $report = $report + "<h4>Leaderboard 'Billed Hours for all Statuses' for this Week as of Today</h4>"      
               
        $sql = $sql + "`r`nSELECT top 100 resource, count(resource) as 'Closed Tickets',sum(Hours_bill) as 'Hours Billed' 
FROM [cwwebapp_nct].[dbo].[v_company_time] where (( date>=DATEADD(week, DATEDIFF(week, 0, getdate()), 0) )) group by resource order by sum(hours_bill) desc;"



        $sql = $sql + "END`r`n"
        #write-host $sql
        # write-host $sql
        $Command.CommandText = $sql
        $datareader = $Command.ExecuteReader()
        IF ($DATAREADER.HASROWS) {
            #write-host "Has rows"



            $report = $report + "<table style=`"width:100%`" class=`"container`">
<tr>




<th style=`"width:20%`">&#9734; Ranking &#9734;</th>
<th>Resource</th>
<th style=`"width:20%`">All Tickets</th>
<th style=`"width:20%`">Hours Billed</th>


</tr>"



            $place = 0
            WHILE ($DATAREADER.READ()) {
        
                $place = $place + 1

                $report = $report + "<tr><td>" + $place + "</td><td>" + $DATAREADER[0] + "</td><td>" + $DATAREADER[1] + "</td><td>" + $DATAREADER[2] + "</td></tr>" 
       
            }

            $report = $report + "</table>"
        }
        $datareader.close()





        #BILLED HOURS week END

        #Billed hours month start

        $sql = "begin`r`n"

        $report = $report + "<h4>Leaderboard 'Billed Hours for all Statuses' for this Month as of Today</h4>"      
               
        $sql = $sql + "`r`nSELECT top 100 resource, count(resource) as 'Closed Tickets',sum(Hours_bill) as 'Hours Billed' 
FROM [cwwebapp_nct].[dbo].[v_company_time] where (( date>=DATEADD(month, DATEDIFF(month, 0, getdate()), 0) )) group by resource order by sum(hours_bill) desc;"



        $sql = $sql + "END`r`n"
        #write-host $sql
        # write-host $sql
        $Command.CommandText = $sql
        $datareader = $Command.ExecuteReader()
        IF ($DATAREADER.HASROWS) {
            #write-host "Has rows"



            $report = $report + "<table style=`"width:100%`" class=`"container`">
<tr>




<th style=`"width:20%`">&#9734; Ranking &#9734;</th>
<th>Resource</th>
<th style=`"width:20%`">All Tickets</th>
<th style=`"width:20%`">Hours Billed</th>


</tr>"



            $place = 0
            WHILE ($DATAREADER.READ()) {
        
                $place = $place + 1

                $report = $report + "<tr><td>" + $place + "</td><td>" + $DATAREADER[0] + "</td><td>" + $DATAREADER[1] + "</td><td>" + $DATAREADER[2] + "</td></tr>" 
       
            }

            $report = $report + "</table>"
        }
        $datareader.close()

        #billed hours month end


        #BILLED HOURS year

        $sql = "begin`r`n"

        $report = $report + "<h4>Leaderboard 'Billed Hours for all Statuses' for this Year as of Today</h4>"      
               
        $sql = $sql + "`r`nSELECT top 100 resource, count(resource) as 'Closed Tickets',sum(Hours_bill) as 'Hours Billed' 
FROM [cwwebapp_nct].[dbo].[v_company_time] where (( date>=DATEADD(year, DATEDIFF(year, 0, getdate()), 0) )) group by resource order by sum(hours_bill) desc;"



        $sql = $sql + "END`r`n"
        #write-host $sql
        # write-host $sql
        $Command.CommandText = $sql
        $datareader = $Command.ExecuteReader()
        IF ($DATAREADER.HASROWS) {
            #write-host "Has rows"



            $report = $report + "<table style=`"width:100%`" class=`"container`">
<tr>




<th style=`"width:20%`">&#9734; Ranking &#9734;</th>
<th>Resource</th>
<th style=`"width:20%`">All Tickets</th>
<th style=`"width:20%`">Hours Billed</th>


</tr>"



            $place = 0
            WHILE ($DATAREADER.READ()) {
        
                $place = $place + 1

                $report = $report + "<tr><td>" + $place + "</td><td>" + $DATAREADER[0] + "</td><td>" + $DATAREADER[1] + "</td><td>" + $DATAREADER[2] + "</td></tr>" 
       
            }

            $report = $report + "</table>"
        }
        $datareader.close()





        #BILLED HOURS year END

        #detail week

        $sql = "begin`r`n"

        $report = $report + "<h4>Leaderboard 'Detail for all Statuses' for this Week as of Today</h4>"      
               
        $sql = $sql + "`r`nselect top 100 Member_Name,sum(datalength(notes))
FROM [cwwebapp_nct].[dbo].[v_TimeRecords] 
where (( date_start>=DATEADD(week, DATEDIFF(week, 0, getdate()), 0) )) group by Member_Name order by sum(datalength(notes)) desc;"



        $sql = $sql + "END`r`n"
        #write-host $sql
        # write-host $sql
        $Command.CommandText = $sql
        $datareader = $Command.ExecuteReader()
        IF ($DATAREADER.HASROWS) {
            #write-host "Has rows"



            $report = $report + "<table style=`"width:100%`" class=`"container`">
<tr>




<th style=`"width:20%`">&#9734; Ranking &#9734;</th>
<th>Resource</th>
<th style=`"width:20%`">Characters</th> 


</tr>"



            $place = 0
            WHILE ($DATAREADER.READ()) {
        
                $place = $place + 1

                $report = $report + "<tr><td>" + $place + "</td><td>" + $DATAREADER[0] + "</td><td>" + $DATAREADER[1] + "</td></tr>" 
       
            }

            $report = $report + "</table>"
        }
        $datareader.close()





        #detail week END





        #detail month

        $sql = "begin`r`n"

        $report = $report + "<h4>Leaderboard 'Detail for all Statuses' for this Month as of Today</h4>"      
               
        $sql = $sql + "`r`nselect top 100 Member_Name,sum(datalength(notes))
FROM [cwwebapp_nct].[dbo].[v_TimeRecords] 
where (( date_start>=DATEADD(month, DATEDIFF(month, 0, getdate()), 0) )) group by Member_Name order by sum(datalength(notes)) desc;"



        $sql = $sql + "END`r`n"
        #write-host $sql
        # write-host $sql
        $Command.CommandText = $sql
        $datareader = $Command.ExecuteReader()
        IF ($DATAREADER.HASROWS) {
            #write-host "Has rows"



            $report = $report + "<table style=`"width:100%`" class=`"container`">
<tr>




<th style=`"width:20%`">&#9734; Ranking &#9734;</th>
<th>Resource</th>
<th style=`"width:20%`">Characters</th> 


</tr>"



            $place = 0
            WHILE ($DATAREADER.READ()) {
        
                $place = $place + 1

                $report = $report + "<tr><td>" + $place + "</td><td>" + $DATAREADER[0] + "</td><td>" + $DATAREADER[1] + "</td></tr>" 
       
            }

            $report = $report + "</table>"
        }
        $datareader.close()





        #detail month END



        #detail year

        $sql = "begin`r`n"

        $report = $report + "<h4>Leaderboard 'Detail for all Statuses' for this Year as of Today</h4>"      
               
        $sql = $sql + "`r`nselect top 100 Member_Name,sum(datalength(notes))
FROM [cwwebapp_nct].[dbo].[v_TimeRecords] 
where (( date_start>=DATEADD(year, DATEDIFF(year, 0, getdate()), 0) )) group by Member_Name order by sum(datalength(notes)) desc;"



        $sql = $sql + "END`r`n"
        #write-host $sql
        # write-host $sql
        $Command.CommandText = $sql
        $datareader = $Command.ExecuteReader()
        IF ($DATAREADER.HASROWS) {
            #write-host "Has rows"



            $report = $report + "<table style=`"width:100%`" class=`"container`">
<tr>




<th style=`"width:20%`">&#9734; Ranking &#9734;</th>
<th>Resource</th>
<th style=`"width:20%`">Characters</th> 


</tr>"



            $place = 0
            WHILE ($DATAREADER.READ()) {
        
                $place = $place + 1

                $report = $report + "<tr><td>" + $place + "</td><td>" + $DATAREADER[0] + "</td><td>" + $DATAREADER[1] + "</td></tr>" 
       
            }

            $report = $report + "</table></body></html>"
        }
        $datareader.close()





        #detail year END













    }
    finally {
        $Connection.Close()
        write-host "Query Complete" 
    }

    $message = new-object Net.Mail.MailMessage;
    
    $message.From = $smtpfrom;
    $message.To.Add($smtpto);
    $message.Subject = $strcomputer + " Leaderboard Report " + (get-date) ;
    $message.IsBodyHTML = $true
    #$message.Headers=$Header
    $message.Body = $report
    
    #$attachment = New-Object Net.Mail.Attachment($attachmentpath);
    #$message.Attachments.Add($attachment);
    set-Content -Path c:\support\leaderboard.html -Value $report
    $smtp = new-object Net.Mail.SmtpClient($smtpserver, $smtpport);
    $smtp.EnableSSL = $true;
    $smtp.Credentials = New-Object System.Net.NetworkCredential($sendUsername, $sendPassword);
    $smtp.send($message);
    write-host "Mail Sent to "  $smtpto ; 
    #$attachment.Dispose();




}
catch {
    $ErrorMessage = $_.Exception.Message
    $FailedItem = $_.Exception.ItemName
    write-host "Error with Item $FailedItem. The error message was $ErrorMessage"
    #Send-MailMessage -From ExpensesBot@MyCompany.Com -To WinAdmin@MyCompany.Com -Subject "HR File Read Failed!" -SmtpServer EXCH01.AD.MyCompany.Com -Body "We failed to read file $FailedItem. The error message was $ErrorMessage"
    Break

}
