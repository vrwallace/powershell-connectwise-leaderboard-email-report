####################################
# Program: leaderboardEMAILreport.ps1
# By: Von Wallace vonwallace@vonwallace.com
# To run add the following to the login script
# powershell.exe –Noninteractive –Noprofile –Command "C:\support\leaderboardEMAILreport.ps1"
# Program queries CW SQL server
###################################



try {

    
   
    #settings
    $Version = "1.01ps"
    $smtpserver = "smtp.office365.com"
    $smtpport = "587"
    $smtpfrom = "someone@somewhere.net"
    
    $smtpto = "someone@somewhere.net"
    
    $sendusername = "someone@somewhere.net"
    $sendpassword = "password"
    
    $uid = "SQLuser"
    $pwd = "sqlpassword"
       
    $sqlserver = "127.0.0.1"
    $SQLDBName = "cwdatabase"

    $Connection = New-Object System.Data.SQLClient.SQLConnection
    $Connection.ConnectionString = "Server = $SQLServer; Database = $SQLDBName; User ID = $uid; Password = $pwd;"
    #$Connection.ConnectionString = "Server = $SQLServer; Database = $SQLDBName; Integrated Security = True;"
    
    $Connection.Open()
    try {
        $Command = New-Object System.Data.SQLClient.SQLCommand
        $Command.Connection = $Connection

        
        $sql = "begin`r`n"

        $report = @"
<style>
TABLE {border-width: 1px; border-style: solid; border-color: black; border-collapse: collapse;}
TH {border-width: 1px; padding: 3px; border-style: solid; border-color: black; background-color: #6495ED;}
TD {border-width: 1px; padding: 3px; border-style: solid; border-color: black;}
</style>
"@    
        $report = $report + "<h4>Leader Board Report</h4>"
        $report = $report + "<h4>Version: " + $version + "</h4><br>"

       
        $report = $report + "<h4>Leader Board 'Closed Tickets' for this Week as of Today</h4>"      
               
        $sql = $sql + "`r`nSELECT top 100 resource, count(resource) as 'Closed Tickets',sum(Hours_bill) as 'Hours Billed' 
        FROM ["+$SQLDBName+"].[dbo].[v_company_time] where (( sr_status='Closed') and ( date>=DATEADD(week, DATEDIFF(week, 0, getdate()), 0) ) ) group by resource order by count(resource) desc;
        "


        $sql = $sql + "END`r`n"
        
        $Command.CommandText = $sql
        $datareader = $Command.ExecuteReader()
        IF ($DATAREADER.HASROWS) {
        


            $report = $report + "<table style=""width:100%"">
  <tr>
    
  
  <th>Ranking</th> 
    <th>Resource</th>
    <th>Closed Tickets</th> 
    <th>Hours Billed</th>
  
    
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

        $report = $report + "<h4>Leader Board 'Closed Tickets' for this Month as of Today</h4>"      
               
        $sql = $sql + "`r`nSELECT top 100 resource, count(resource) as 'Closed Tickets',sum(Hours_bill) as 'Hours Billed' 
FROM ["+$SQLDBName+"].[dbo].[v_company_time] where (( sr_status='Closed') and ( date>=DATEADD(month, DATEDIFF(month, 0, getdate()), 0) )) group by resource order by count(resource) desc;"



        $sql = $sql + "END`r`n"
        $Command.CommandText = $sql
        $datareader = $Command.ExecuteReader()
        IF ($DATAREADER.HASROWS) {



            $report = $report + "<table style=""width:100%"">
<tr>


<th>Ranking</th>
<th>Resource</th>
<th>Closed Tickets</th> 
<th>Hours Billed</th>


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
        $report = $report + "<h4>Leader Board 'Closed Tickets' for this Year as of Today</h4>"      
               
        $sql = $sql + "`r`nSELECT top 100 resource, count(resource) as 'Closed Tickets',sum(Hours_bill) as 'Hours Billed' 
FROM ["+$SQLDBName+"].[dbo].[v_company_time] where (( sr_status='Closed') and ( date>=DATEADD(year, DATEDIFF(year, 0, getdate()), 0) )) group by resource order by count(resource) desc;"



        $sql = $sql + "END`r`n"
        $Command.CommandText = $sql
        $datareader = $Command.ExecuteReader()
        IF ($DATAREADER.HASROWS) {



            $report = $report + "<table style=""width:100%"">
<tr>


<th>Ranking</th>
<th>Resource</th>
<th>Closed Tickets</th> 
<th>Hours Billed</th>

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

        #All time
        $sql = "begin`r`n"
        $report = $report + "<h4>Leader Board 'Closed Tickets' All Time as of Today</h4>"      
               
        $sql = $sql + "`r`nSELECT top 100 resource, count(resource) as 'Closed Tickets',sum(Hours_bill) as 'Hours Billed' 
FROM ["+$SQLDBName+"].[dbo].[v_company_time] where (( sr_status='Closed')) group by resource order by count(resource) desc;"



        $sql = $sql + "END`r`n"
        $Command.CommandText = $sql
        $datareader = $Command.ExecuteReader()
        IF ($DATAREADER.HASROWS) {


            
            $report = $report + "<table style=""width:100%"">
<tr>


<th>Ranking</th>
<th>Resource</th>
<th>Closed Tickets</th> 
<th>Hours Billed</th>


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

        $report = $report + "<h4>Leader Board 'Billed Hours for all Statuses' for this Week as of Today</h4>"      
               
        $sql = $sql + "`r`nSELECT top 100 resource, count(resource) as 'Closed Tickets',sum(Hours_bill) as 'Hours Billed' 
FROM ["+$SQLDBName+"].[dbo].[v_company_time] where (( date>=DATEADD(week, DATEDIFF(week, 0, getdate()), 0) )) group by resource order by sum(hours_bill) desc;"



        $sql = $sql + "END`r`n"
        $Command.CommandText = $sql
        $datareader = $Command.ExecuteReader()
        IF ($DATAREADER.HASROWS) {



            $report = $report + "<table style=""width:100%"">
<tr>


<th>Ranking</th>
<th>Resource</th>
<th>All Tickets</th> 
<th>Hours Billed</th>


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

        $report = $report + "<h4>Leader Board 'Billed Hours for all Statuses' for this Month as of Today</h4>"      
               
        $sql = $sql + "`r`nSELECT top 100 resource, count(resource) as 'Closed Tickets',sum(Hours_bill) as 'Hours Billed' 
FROM ["+$SQLDBName+"].[dbo].[v_company_time] where (( date>=DATEADD(month, DATEDIFF(month, 0, getdate()), 0) )) group by resource order by sum(hours_bill) desc;"



        $sql = $sql + "END`r`n"
        $Command.CommandText = $sql
        $datareader = $Command.ExecuteReader()
        IF ($DATAREADER.HASROWS) {

            $report = $report + "<table style=""width:100%"">
<tr>


<th>Ranking</th>
<th>Resource</th>
<th>All Tickets</th> 
<th>Hours Billed</th>


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

        $report = $report + "<h4>Leader Board 'Billed Hours for all Statuses' for this Year as of Today</h4>"      
               
        $sql = $sql + "`r`nSELECT top 100 resource, count(resource) as 'Closed Tickets',sum(Hours_bill) as 'Hours Billed' 
FROM ["+$SQLDBName+"].[dbo].[v_company_time] where (( date>=DATEADD(year, DATEDIFF(year, 0, getdate()), 0) )) group by resource order by sum(hours_bill) desc;"



        $sql = $sql + "END`r`n"
        $Command.CommandText = $sql
        $datareader = $Command.ExecuteReader()
        IF ($DATAREADER.HASROWS) {

            $report = $report + "<table style=""width:100%"">
<tr>


<th>Ranking</th>
<th>Resource</th>
<th>All Tickets</th> 
<th>Hours Billed</th>


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

        $report = $report + "<h4>Leader Board 'Detail for all Statuses' for this Week as of Today</h4>"      
               
        $sql = $sql + "`r`nselect top 100 Member_Name,sum(datalength(notes))
FROM ["+$SQLDBName+"].[dbo].[v_TimeRecords] 
where (( date_start>=DATEADD(week, DATEDIFF(week, 0, getdate()), 0) )and (Member_Name <> 'Paul Indarawis')) group by Member_Name order by sum(datalength(notes)) desc;"



        $sql = $sql + "END`r`n"
        $Command.CommandText = $sql
        $datareader = $Command.ExecuteReader()
        IF ($DATAREADER.HASROWS) {

            $report = $report + "<table style=""width:100%"">
<tr>


<th>Ranking</th>
<th>Resource</th>
<th>Characters</th> 


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

        $report = $report + "<h4>Leader Board 'Detail for all Statuses' for this Month as of Today</h4>"      
               
        $sql = $sql + "`r`nselect top 100 Member_Name,sum(datalength(notes))
FROM ["+$SQLDBName+"].[dbo].[v_TimeRecords] 
where (( date_start>=DATEADD(month, DATEDIFF(month, 0, getdate()), 0) )and (Member_Name <> 'Paul Indarawis')) group by Member_Name order by sum(datalength(notes)) desc;"



        $sql = $sql + "END`r`n"
        $Command.CommandText = $sql
        $datareader = $Command.ExecuteReader()
        IF ($DATAREADER.HASROWS) {

            $report = $report + "<table style=""width:100%"">
<tr>


<th>Ranking</th>
<th>Resource</th>
<th>Characters</th> 

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

        $report = $report + "<h4>Leader Board 'Detail for all Statuses' for this Year as of Today</h4>"      
               
        $sql = $sql + "`r`nselect top 100 Member_Name,sum(datalength(notes))
FROM ["+$SQLDBName+"].[dbo].[v_TimeRecords] 
where (( date_start>=DATEADD(year, DATEDIFF(year, 0, getdate()), 0) )and (Member_Name <> 'Paul Indarawis')) group by Member_Name order by sum(datalength(notes)) desc;"



        $sql = $sql + "END`r`n"
        $Command.CommandText = $sql
        $datareader = $Command.ExecuteReader()
        IF ($DATAREADER.HASROWS) {

            $report = $report + "<table style=""width:100%"">
<tr>

<th>Ranking</th>
<th>Resource</th>
<th>Characters</th> 

</tr>"
            $place = 0
            WHILE ($DATAREADER.READ()) {
        
                $place = $place + 1

                $report = $report + "<tr><td>" + $place + "</td><td>" + $DATAREADER[0] + "</td><td>" + $DATAREADER[1] + "</td></tr>" 
       
            }

            $report = $report + "</table>"
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
    $message.Subject = $strcomputer + " Leader Board Report " + (get-date) ;
    $message.IsBodyHTML = $true
    $message.Body = $report
    

    $smtp = new-object Net.Mail.SmtpClient($smtpserver, $smtpport);
    $smtp.EnableSSL = $true;
    $smtp.Credentials = New-Object System.Net.NetworkCredential($sendUsername, $sendPassword);
    $smtp.send($message);
    write-host "Mail Sent to "  $smtpto ; 


}
catch {
    $ErrorMessage = $_.Exception.Message
    $FailedItem = $_.Exception.ItemName
    write-host "Error with Item $FailedItem. The error message was $ErrorMessage"
    Break

}
