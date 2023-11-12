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
    
    #$smtpfrom = "nctsqlalerts@nctnet.net"

    $smtpfrom = "leaderboard@nctnet.net"
    
    $smtpto = "n@nctnet.net"
    
    #    $sendusername = "nctsqlalerts@nctnet.net"
    #    $sendpassword =  "bwktfvnxmskdhqyt"

    $sendusername = "leaderboard@nctnet.net"
    $sendpassword = "ul+raHeat46ul+raHeat46"
    
    $uid = "SQLAlerts"
    $pwd = "w!ndyTurtle10"
       
    $sqlserver = "10.151.1.157"
    $SQLDBName = "cwwebapp_nct"

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
        $report = $report + "<center><img src=`"data:image/jpeg;base64,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`" width=`"203`" height=`"91`"/>"
        $report = $report + "<br/><br/><br/>"

        $report = $report + "<h1>Leaderboard Report</h1>"
        $report = $report + "<h4>Version: " + $version + "</h4></center><br>"

       
        $report = $report + "<h4>Leaderboard 'Closed Tickets' for this Week as of Today</h4>"      
               
        $sql = $sql + "`r`nSELECT top 100 resource, count(resource) as 'Closed Tickets',sum(Hours_bill) as 'Hours Billed' 
        FROM [cwwebapp_nct].[dbo].[v_company_time] where (( sr_status='Closed') and ( date>=DATEADD(week, DATEDIFF(week, 0, getdate()), 0) ) and (resource <> 'Paul Indarawis')) group by resource order by count(resource) desc;
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
FROM [cwwebapp_nct].[dbo].[v_company_time] where (( sr_status='Closed') and ( date>=DATEADD(month, DATEDIFF(month, 0, getdate()), 0) )and (resource <> 'Paul Indarawis')) group by resource order by count(resource) desc;"



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
FROM [cwwebapp_nct].[dbo].[v_company_time] where (( sr_status='Closed') and ( date>=DATEADD(year, DATEDIFF(year, 0, getdate()), 0) )and (resource <> 'Paul Indarawis')) group by resource order by count(resource) desc;"



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
FROM [cwwebapp_nct].[dbo].[v_company_time] where (( date>=DATEADD(week, DATEDIFF(week, 0, getdate()), 0) )and (resource <> 'Paul Indarawis')) group by resource order by sum(hours_bill) desc;"



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
FROM [cwwebapp_nct].[dbo].[v_company_time] where (( date>=DATEADD(month, DATEDIFF(month, 0, getdate()), 0) )and (resource <> 'Paul Indarawis')) group by resource order by sum(hours_bill) desc;"



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
FROM [cwwebapp_nct].[dbo].[v_company_time] where (( date>=DATEADD(year, DATEDIFF(year, 0, getdate()), 0) )and (resource <> 'Paul Indarawis')) group by resource order by sum(hours_bill) desc;"



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
where (( date_start>=DATEADD(week, DATEDIFF(week, 0, getdate()), 0) )and (Member_Name <> 'Paul Indarawis')) group by Member_Name order by sum(datalength(notes)) desc;"



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
where (( date_start>=DATEADD(month, DATEDIFF(month, 0, getdate()), 0) )and (Member_Name <> 'Paul Indarawis')) group by Member_Name order by sum(datalength(notes)) desc;"



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
where (( date_start>=DATEADD(year, DATEDIFF(year, 0, getdate()), 0) )and (Member_Name <> 'Paul Indarawis')) group by Member_Name order by sum(datalength(notes)) desc;"



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
