#Load\Save Variables for alteration.
$path = 'C:\Users\pete\Desktop\users.csv'
$savepath = 'C:\Users\pete\Desktop\site-summary.csv'

#QUESTION1
echo "Q1:Import Users.csv as a $"

$results = Import-Csv -Path $path
echo "Importing from $path, if file not found change the 'path' variable"

#NOTE: If you had this data split across a number of csvs in the same data format, you could store these csvs in the same folder
#location and could import them all into the same variable using the following code:

#set folder location (but not file or extension)
#$csvfolder = 'C:\Users\pete\Desktop\'
#indentify csv files in folder and 
#$results = Import-Csv -Path  (Get-ChildItem -Path $csvfolder -Filter '*.csv').FullName
 
#QUESTION2
echo "Q2:How Many Users are there?" #A:148
$results.count

#QUESTION3
echo "Q3:What is the total size of all mailboxes?" #A:3936.56gb
($results.mailboxsizegb | measure-object -sum).sum

#QUESTION4
echo "Q4:How many accounts exist with non-identical EmailAddress/UserPrincipalName?" #Be mindful of case sensitivity. A: 143

#The first line only outputs the email address into the new variable if UPN = Email Address
#Needs to use -ceq to be case sensitive 
$Q3 = foreach ($x in $results) {if ($x.UserPrincipalName -ceq $x.EmailAddress) {$x.EmailAddress}}
$Q3.count

#QUESTION5
echo "Q5:Same as question 3 (total mailbox size), but limited only to Site = NYC?" # A:938.9

#Extract only the NYC = site objects in a new variable
$OnlyNYC = $results | Where-Object {$_.site -eq "NYC"}
#use measure-object -sum to find the total
($OnlyNYC.mailboxsizegb | measure-object -sum).sum

#NOTE: you can do this in one line as well, without the need for an extra variable but I don't think it is as clear 
#(to use this un-rem the below)
#($results | Where-Object {$_.site -eq "NYC"}).mailboxsizegb | measure-object -sum

#QUESTION 6
echo "Q6: How many Employees (AccountType: Employee) have mailboxes larger than 10 GB?" #(remember MailboxSizeGB is already in GB.)

$not10 = ($results | Where-Object {$_.AccountType -eq "Employee" -and [int]$_.mailboxsizegb -gt 10})
$not10.count

#or on one line 
#($results | Where-Object {$_.AccountType -eq "Employee" -and [int]$_.mailboxsizegb -gt 10}).count

#I like the above question as the casting of the mailbox size as a [int] is important - without it not all system objects are recognised
#as integers, so some are not removed.
#THIS will not give the correct answer: ($results | Where-Object {$_.AccountType -eq "Employee" -and $_.mailboxsizegb -gt 10}).count
 
 
#QUESTION 7

echo "Q7: Provide a list of the top 10 users with EmailAddress @domain2.com in Site: NYC by mailbox size, descending."
#a.	The boss already knows that they’re @domain2.com; he wants to only know their usernames, that is, the part
#of the EmailAddress before the “@” symbol.  There is suspicion that IT Admins managing domain2.com
# are a quirky bunch and are encoding hidden messages in their directory via email addresses.  
#Parse out these usernames (in the expected order) and place them in a single string, separated by spaces –
#should look like: “user1 user2 … user10”

#A: ok so it appears that i know what im doing

$Q7part1 = ($results | Where-Object {$_.site -eq "NYC" -and $_.EmailAddress -match "@domain2.com"})
$Q7part2 = $Q7part1 | sort-object mailboxsizegb -Descending | select-object -First 10
$Q7part3 = ($Q7part2.emailaddress) -replace "@domain2.com"
$Q7part4 = $Q7part3 -join " "
$Q7part4

#QUESTION 8

echo "Q8: Create a new CSV file that summarizes Sites, using the following headers: 
#Site, TotalUserCount, EmployeeCount, ContractorCount, TotalMailboxSizeGB, AverageMailboxSizeGB"
#Create this CSV file based off of the original Users.csv.  
#Note that the boss is picky when it comes to formatting – make sure that AverageMailboxSizeGB is formatted
#to the nearest tenth of a GB (e.g. 50.124124 is formatted as 50.1).  
#You must use PowerShell to format this because Excel is down for maintenance.

$results = Import-Csv -Path $path

#Groups the results together in sites
$groupedresults = $results | group site

#For each site, add the $sites properties to a new PScustomobject to get the formatting neccessary.
$Q8 = foreach ($site in $groupedresults) {
 
           [PSCustomObject]@{
            Site = $site.name
            TotalUserCount = $site.count
            EmployeeCount = ($site.group.accounttype | Select-string "Employee").count
            ContactorCount = ($site.group.accounttype | Select-string "Contractor").count
            #These last two properties also contain '“{0:f1}” -f' formatting to ensure formatted to 1 decimal place.
            TotalMailboxSizeGB = (“{0:f1}” -f ($site.group.mailboxsizegb | measure-object -sum).sum)
            AverageMailboxSizeGB = (“{0:f1}” -f ($site.group.mailboxsizegb | measure-object -Average).average)
    }
}
#Once built export the custom object to csv. Change the $savepath variable to change the file location.

#$savepath = 'C:\Users\pete\Desktop\site-summary.csv' (listed at top of powershell file)
$q8 | ft
echo "Saving csv to $savepath, if error check the 'savepath' variable is set correctly"
$q8 | export-Csv -Path $savepath -NoTypeInformation

#Results =
#Site TotalUserCount EmployeeCount ContactorCount TotalMailboxSizeGB AverageMailboxSizeGB
#---- -------------- ------------- -------------- ------------------ --------------------
#TOR              29            15             14 514.3              17.7                
#SEA              42            21             21 1092.0             26.0                
#BOS               4             3              1 370.0              92.5                
#LAS              21            10             11 343.1              16.3                
#BRZ              21            11             10 340.1              16.2                
#RIO              21            10             11 338.1              16.1                
#NYC              10             8              2 938.9              93.9      

echo "Pause in case not using Powershell ISE to run"
Read-Host -Prompt "Press Enter to exit"          