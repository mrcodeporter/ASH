
 
 #$emailslist
 
 # Automated Sent Bulk Email v0.5
# DATE: 2022 
# VERSION: 0.5
<#
    
 Automated Sent Bulk Email Alpha 0.5:

    Tools for administrative  Sent out bulk emails to Terminations users 
    This script was created with the help of a lot of google stuff, and
    Reading Learn PowerShell in a Month of Lunches is an excellent book.


  Required Software:

    PowerShell v5 or higher

    
 Automated Sent Bulk EMail Alpha 0.5
    work on this ervin 
    Tools for administrative Sent out bulk emails to Terminations users 
    This script was created with the help of a lot of google stuff,and
    Reading Learn PowerShell Scripting in a Month of Lunches is an excellent book.

 Tips:
    - work on this ervin 
	- For best results this script should be run with elevated privileges
	- 
	TODO:
	- Alot :P work on this 



#>


#----------------------------------------------------------------------------------------#
#-- Header 
#----------------------------------------------------------------------------------------#

function Title_header 
 {
   Clear-Host;
   Write-Host "-----------------------------------------";
   Write-Host -BackgroundColor Magenta "        AUTOMATED SENT BULK EMAILS      "; 
   Write-Host -BackgroundColor Blue "                Alpha 0.5               ";
   #Write-Host "                                         ";
   Write-Host  " Check notes in Script 
        for more additional information.";
   Write-Host "-----------------------------------------";
   
   
 };



#-- Main Menu Header 
#----------------------------------------------------------------------------------------#

 function Main_Menu_header 
 {
Write-Host "======================================="
Write-Host "====" -noNewline
Write-Host "=========" -noNewline
write-host " MAIN MENU " -foregroundcolor black -backgroundcolor yellow -noNewline
write-host "==============="
Write-Host "======================================="
Write-Host -ForegroundColor DarkYellow  "  Please Choose One  
          Of The Options Below: ";
		

    Write-Host "`n";

		Write-Host "-[1] Files Located ";
        Write-Host "`n";
		Write-Host "-[2] Users ";
        Write-Host "`n";
		Write-Host "-[3] Compose Email ";
		Write-Host "`n";

Write-Host -backgroundcolor blue "-[] More Features Coming Soon!"; 
		Write-Host "-[Q] Quit";
		Write-Host "`n";
 }




 #-- Menu Files Located 
#----------------------------------------------------------------------------------------#
 function Files_Located_Header
 {
  

    
   
  Write-Host "======================================"
  Write-Host "====" -noNewline
  Write-Host "=========" -noNewline
  write-host " FILES LOCATED " -foregroundcolor black -backgroundcolor yellow -noNewline
  write-host "=========="
  Write-Host "======================================"
  Write-Host -ForegroundColor DarkYellow "  What would you like to do?: ";
      

  Write-Host "`n";
  Write-Host "-[1] Upload CSV File ";
  Write-Host "`n";
  Write-Host "-[2] Create Temp CSV File ";
  Write-Host "`n";
  Write-Host "-[3] Back to Main Menu ";
  Write-Host "`n";

  Write-Host -backgroundcolor blue "-[] More Features Coming Soon!"; 
  Write-Host "-[Q] Quit";
  Write-Host "`n";
   
 }


#-- Users Menu 
#----------------------------------------------------------------------------------------#
 function Users_Main_Menu_Header
 {
  
   Clear-Host;
  &Title_header;
   
  Write-Host "======================================"
  Write-Host "====" -noNewline
  Write-Host "=========" -noNewline
  write-host " Users Menu " -foregroundcolor black -backgroundcolor yellow -noNewline
  write-host "=========="
  Write-Host "======================================"
  Write-Host -ForegroundColor DarkYellow "  What would you like to do?: ";
      

  Write-Host "`n";
  Write-Host "-[1] Add New Users ";
  Write-Host "`n";
  Write-Host "-[2] List of Users  ";
  Write-Host "`n";
  Write-Host "-[3] Back to Main Menu ";
  Write-Host "`n";

  Write-Host -backgroundcolor blue "-[] More Features Coming Soon!"; 
  Write-Host "-[Q] Quit";
  Write-Host "`n";
   
 }

#----------------------------------------------------------------------------------------#
#-- End  this all work so far


#----------------------------------------------------------------------------------------#


  

#----------------------------------------------------------------------------------------#
#-- Files Location.
#----------------------------------------------------------------------------------------#

function Open-file([string]$initialditrctory) {

  [System.Reflection.Assembly]::LoadWithPartialName("system.windows.forms") | Out-Null

  $openFileDialog = New-Object System.Windows.Forms.OpenFileDialog 
  $openFileDialog.InitialDirectory = $currentDirectory 
  $openFileDialog.Filter = "CSV (*.csv)| *.csv" 
  $openFileDialog.ShowDialog() | Out-Null 
  $openFileDialog.Filename 






}

function Save-File([string] $initialDirectory ) 

{

[System.Reflection.Assembly]::LoadWithPartialName("System.windows.forms") | Out-Null

$OpenFileDialog = New-Object System.Windows.Forms.SaveFileDialog
$OpenFileDialog.initialDirectory = $initialDirectory
$OpenFileDialog.FileName = "Temporary files"
$OpenFileDialog.filter = "CSV (*.csv)| *.csv"
$OpenFileDialog.ShowDialog() |  Out-Null

return $OpenFileDialog.FileName  

} 




#----------------------------------------------------------------------------------------#
#-- MENUS
#----------------------------------------------------------------------------------------#


function Main_Menu
 {
      Clear-Host;
   
   &Title_header;
   &Main_Menu_header;
    
     do {
 
$user_input= Read-Host "Enter Choice"

$Correct_Array=@()
$Correct_Array= 1..5 

$exit_Array=@()
$exit_Array= "q","quit"

$turnString= [string]$Correct_Array

switch ($user_input)
  {
      1 { Files_Located_Menu } 
      2 { Users_Main_Menu } 
      3 {Sent_Email}
      4 {"get_email"}
      q {     "Build a exit screen with a funtion"           }
      quit {     "Build a exit screen with a funtion"           }
  }

} until (

  ( $user_input -eq '1' ) -or  ( $user_input -eq '2' )  -or  ( $user_input -eq '3' ) -or ( $user_input -eq '4' )  -or ( $user_input -eq '5' )  -or ($user_input -eq "q") -or ( $user_input -eq 'quit' ) 
  
)

 }



 function Files_Located_Menu
{
  

  &Title_header;
  &Files_Located_Header;
  
  do {
  
    
$user_input= Read-Host "Enter Choice"

$whotest=@()
$whotest= 1..5

$turnString= [string]$whotest 

$exit_Array=@()
$exit_Array= 'q','quit'

switch ($user_input)
  {
      1 { $open_mefile =Open-file

if ( $open_mefile -ne "" ) 
{
Write-Output "You Choose Open File.: $open_mefile" 
 
} 

else 
{
Write-Output " No File Was Chosen."
}   } 
      
      
       
      2 {$File= Save-File 

if ( $File -ne "Temporary files" ) 
{
Write-Output "You choose File Path: $File" 

 #Create an array that also stores details of all the information that need to be grabbed. 
  $info = @()
  $info = '"First Name","Last Name","Legal / Preferred Address: Address Line 1","Legal / Preferred Address: Address Line 2","Legal / Preferred Address: City","Legal / Preferred Address: State / Territory Code","Primary Address: Zip / Postal Code","Email","Work Contact: Work Email","Personal Contact: Personal Mobile"'


} 

else 
{
Write-Output "No File Path was chosen."
}


  }  
      3 { Main_Menu }      
      4 {"`nIt is four."}
      5 {"`testemail"}
  }
  
 
if (($whotest -match $user_input ) -or ($exit_Array -match $user_input)) {
  
  
   
 
}
else {

  write-host "Sorry that is not one of the choices try again"
 
}



} until (

   ( $user_input -eq '3' ) -or ( $user_input -eq '4' )  -or ( $user_input -eq '5' ) -or ($user_input -eq "q") -or ( $user_input -eq 'quit' )


)


}

function Users_CSV_File {


   Clear-Host;
   
   &Title_header;
   
   #look into add more later
   
   Write-Host "`n";
   Write-Host "`n";
   
   
   
   
   
   Write-host -BackgroundColor Blue "Now I can explore users in the CSV file you upload. Let's begin by typing in their FIRST NAME." 
         
   
   Write-Host "================  FIRST NAME  ================"
   
   $FirstName = Read-Host "Please type in FIRST NAME" 
   $testfile = Import-Csv $open_mefile | Where-Object { ($_."First Name" -eq $FirstName )}
   
   if ($testfile -match $FirstName  ) {
       Write-Host "Found FIRST NAME " -ForegroundColor green   
                             
   }
   else {
     write-host "Cannot find the FIRST NAME " -ForegroundColor Red -WarningAction Inquire
                             
                     
   }
   
        
        {
        
        }
   
   
   Write-Host "================  LAST NAME  ================"
   $LastName = Read-Host "Please type in LAST NAME "
   
   {
   }
   $testfile = Import-Csv $open_mefile | Where-Object { ($_."First Name" -eq $FirstName ) -and ($_."Last Name" -eq $LastName ) }
   if ($testfile -match $LastName  ) {
     Write-Host "Found FIRST and LAST NAME  " -ForegroundColor Green
     Write-Host "I am now Searching for the user." 
      Correct-user 
      Show-CorrectMenu
      
                                 
   }
   else {
     write-host "Cannot find the Last name " -ForegroundColor Red -WarningAction Inquire
     notCorrect-user                
   }
                  
   } 
#----------------------------------------------------------------------------------------#
# :test
#----------------------------------------------------------------------------------------#



function List_Of_Users {

    
   
   &Title_header;

  Write-Host "================ User information ================"
  Import-Csv $File 
  Write-Host "`n";
  option_menu
 


$selection = Read-Host "Please make a selection"
  switch ($selection)
  {
  '1' {
  'You chose option #1'
      Write-Host -ForegroundColor Green 'Add user to list'  'Press anything Key to continue.' 
      pause;
      return Users_Main_Menu
  } 
  
  
  '2' {
  'You chose option #2'
  Write-Host -ForegroundColor red 'No new user was added to the list.' 'Press anything Key to continue.' 
  return Users_CSV_File



  } '3' {
    'You chose option #3'
    return Main_Menu
  }

}   
  
}






#----------------------------------------------------------------------------------------#
# : Menu Users: 
#----------------------------------------------------------------------------------------#

    




      

function Show-CorrectMenu {
  param (
      [string]$Title = 'Choose One?'
  )


$selection = Read-Host "Please make a selection"
  switch ($selection)
  {
  '1' {
  'You chose option #1'
      Write-Host -ForegroundColor Green 'Add user to list'  'Press anything Key to continue.' 
      $testfile | Export-Csv $File  -NoTypeInformation -Force -Append  
      pause;
      return Users_Main_Menu
  } 
  
  
  '2' {
  'You chose option #2'
  Write-Host -ForegroundColor red 'No new user was added to the list.' 'Press anything Key to continue.' 
  return Users_CSV_File



  } '3' {
    'You chose option #3'
    return Main_Menu
  }
  }
  


  }

$File


   function Correct-user
{
  param (
      [string]$Title = 'Choose One?' 


      
  )
  
  Write-Host "================ User information ================"
   $testfile
  

  Write-Host "================ $Title ================"

 
  Write-Host "1: Press '1' if Users is correct"
        
  Write-Host "2: Press '2' if Users is incorrect."
  
  Write-Host -NoNewline -ForegroundColor Yellow 'Do you want to quit Press [Q] '




  #$Exit= $quit |Write-Host -BackgroundColor DarkMagenta -ForegroundColor DarkYellow
  #Write-Host "3: Press '3' for this option."
  #Write-Host "Q: Press 'Q' to quit."
}











function notCorrect-user {

   

Write-Host -BackgroundColor DarkRed -Verbose "================ Oops! I wasn't able to find $FirstName $LastName ================" 
   
  

#Write-Host -BackgroundColor 

Write-Host "1: Press '1' Would you like to try look for a different user?"

Write-Host "2: Press '2' Move on"

$quit = Write-Host -NoNewline -ForegroundColor Yellow 'Do you want to quit Press [Q]: ' 


#Write-Host -NoNewline -ForegroundColor Yellow 'Create Software Inventory Report? [y|n]: '
#Write-Host -BackgroundColor Red  "3: Press 'Q' to quit."
#Write-Host "3: Press '3' for this option."
#Write-Host "Q: Press 'Q' to quit."



}



function user-CorrectMenu {
  param (
      [string]$Title = 'Choose One?'
  )


$selection = Read-Host "Please make a selection"
  switch ($selection)
  {
  '1' {
  'You chose option #1'
      Write-Host -ForegroundColor Green 'Add user to list'  'Press anything Key to continue.' 
      pause;
      return Users_Main_Menu
  } 
  
  
  '2' {
  'You chose option #2'
  Write-Host -ForegroundColor red 'No new user was added to the list.' 'Press anything Key to continue.' 
  return Users_CSV_File



  } '3' {
    'You chose option #3'
    return Main_Menu
  }
  }
  


  }





   function option_menu
{
  param (
      [string]$Title = 'Choose One?' 


      
  )
  
  Write-Host "================ $Title ================"

 
  Write-Host "1: Press '1' if Users is correct"
        
  Write-Host "2: Press '2' if any Users is incorrect."
  
  Write-Host -NoNewline -ForegroundColor Yellow 'Do you want to quit Press [Q] '




  #$Exit= $quit |Write-Host -BackgroundColor DarkMagenta -ForegroundColor DarkYellow
  #Write-Host "3: Press '3' for this option."
  #Write-Host "Q: Press 'Q' to quit."
}









function Users_Main_Menu
{


    
 
   &Title_header;
 &Users_Main_Menu_Header;

     
   do {


$user_input= Read-Host "Enter Choice"

$Correct_Array=@()
$Correct_Array= 1..5 


$exit_Array=@()
$exit_Array= "q","quit"

$turnString= [string]$Correct_Array



switch ($user_input)
{
    1 { Users_CSV_File } 
    2 {List_Of_Users} 
    3 {Main_Menu}
    4 {""}

    q {     "Build a exit screen with a funtion"           }
    quit {     "Build a exit screen with a funtion"           }
}





} until (

( $user_input -eq '1' ) -or  ( $user_input -eq '2' )  -or  ( $user_input -eq '3' ) -or ( $user_input -eq '4' )  -or ( $user_input -eq '5' )  -or ($user_input -eq "q") -or ( $user_input -eq 'quit' ) 



)



}



$HTMLINFO= @" 
 <!DOCTYPE html>
 <html lang="en">
 <head>
   <meta charset="UTF-8">
   <meta http-equiv="X-UA-Compatible" content="IE=edge">
   <meta name="viewport" content="width=device-width, initial-scale=1.0">
   <title>Email</title>
   <style>
 
 
 
   #logo{
 
   width: 115px;
   height: 50px;
   object-fit: scale-down;
 
   }
 
   #Return{
 
   background-color: rgb(229, 0, 0);
   width: 300px;
   border: 15px solid green;
   padding: 50px;
   margin: 20px;
 
 
   }
 
 
 
   .mainbox{
 
   border: 15px;
   background-color: rgb(ffffff);
   margin: 15px;
   padding-left: 15px;
   box-shadow: 0 3px 10px rgba(250, 122, 76);
   position: absolute;
 
 
   }
 
   .Top-Header1{
 
   color: rgb(255, 255, 255);
   background-color: #0066b2;
   border: 1px solid black;
 
 
 
 
   }
 
 
 
   </style>
 </head>
 <body>
   <div class="mainbox"> 
     <!-- BLoom Logo   -->
     <img id="logo" src="https://lists.office.com/Images/12001b61-2cab-4ffd-8dd1-a5f6c122a7c4/1218b0b1-8da8-4293-93f9-37501cfcf28f/TB6MKBOW1STGYZT2WFDRCPEQPK/99d17588-d602-41cd-9538-0ef43c5b3680" alt="99d17588-d602-41cd-9538-0ef43c5b3680">
     <h3 class="Top-Header1">RETURN OF BLOOM/ADVISE PROPERTY</h3>
     <p id="Top-Header2">Greetings, </p>This is Bloom Insurance's Equipment Returns team. We received notice of your departure from the company and wanted to<br>
     reach out regarding returning your equipment. Could you please provide us with the following:
     <ul>
       <li>Your current address</li>
       <li>Whether you will need a box shipped to you</li>
       <li>If there is any other pertinent information,</li>
       <li>such as days that you won't be home to accept a delivery</li>
     </ul><br>
     Once we have this information, we can get you all the necessary materials for you to return your equipment. You will want to pack up your equipment and use the label we provide to ensure that the package returns to our main facility. Your package will need to be taken to your nearest FedEx drop-off location, which can usually be found by searching 'FedEx drop off near me on Google.
     <p>If you have any questions or concerns, please feel free to reach out, and we will be more than happy to assist you.</p>Thanks,<br>
     WFH Returns Team<br>
     Bloom Insurance Agency<br>
   
 </body>
 </html>
"@


 












function Sent_Email {
 
$newfile = import-csv $File 

$newfile | ForEach-Object {
    $Email = $_.'Personal Contact: Personal Email'
    $Username = $_.Username 


    
}

$outlook= new-object -ComObject outlook.application
    
   $mitem= $outlook.CreateItem("olmailitem")
    
   $mitem.Subject = "Just testing this out"
   
   $mitem.To = "$Email"
   $mitem.HTMLBody = $HTMLINFO
   
   $mitem.Send()
}


#run
Main_Menu
 Read-Host "Exit Now"




 