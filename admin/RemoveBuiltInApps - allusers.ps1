# Removes both provisioned and installed applications

$AppsList = "Microsoft.BingFinance",
"Microsoft.BingNews",
"Microsoft.BingWeather",
"Microsoft.BingTravel",
"Microsoft.BingSports",
"Microsoft.Reader",
"Microsoft.BingHealthAndFitness",
"Microsoft.FoodAndDrink",
"Microsoft.SkypeApp",
"Microsoft.MicrosoftSolitaireCollection",
"Microsoft.ZuneMusic",
"Microsoft.ZuneVideo",
"Microsoft.MicrosoftOfficeHub",
"microsoft.windowscommunicationsapps",
"Microsoft.Office.OneNote",
"Microsoft.officeHub",
"Microsoft.OneConnect",
"Microsoft.Office.Sway",
"Microsoft.ConnectivityStore",
"Microsoft.3DBuilder",
"Microsoft.CommsPhone",
"Microsoft.People",
"Microsoft.WindowsPhone",
"Microsoft.Messaging",
"Microsoft.WindowsPhone",
"Microsoft.Phone",
"Microsoft.Twitter",
"Microsoft.WindowsSoundRecorder",
"Microsoft.WindowsCamera",
"Microsoft.XboxApp",
"Microsoft.Windows.Photos",
"Microsoft.WindowsMaps",
"king.com.CandyCrushSodaSaga",        
"Microsoft.StorePurchaseApp",    
"Microsoft.WindowsFeedbackHub",  
"Microsoft.WindowsStore", 
"Microsoft.XboxIdentityProvider",
#Microsoft.WindowsAlarms 
#Microsoft.AAD.BrokerPlugin               
#Microsoft.AccountsControl                
#Microsoft.BioEnrollment                  
#Microsoft.LockApp                        
#"Microsoft.MicrosoftEdge",
#Microsoft.PPIProjection                  
#Microsoft.Windows.Apprep.ChxApp          
#Microsoft.Windows.AssignedAccessLockApp  
#Microsoft.Windows.CloudExperienceHost    
#Microsoft.Windows.ContentDeliveryManager                 
#Microsoft.Windows.ParentalControls       
#Microsoft.Windows.SecondaryTileExperience
#Microsoft.Windows.SecureAssessmentBrowser
#Microsoft.Windows.ShellExperienceHost    
"Microsoft.XboxGameCallableUI",             
"Windows.ContactSupport",                   
#windows.immersivecontrolpanel            
#Windows.MiracastView                     
#Windows.PrintDialog                      
"Microsoft.Advertising.Xaml",
"Microsoft.Advertising.Xaml",               
#Microsoft.DesktopAppInstaller            
#Microsoft.MicrosoftStickyNotes           
#Microsoft.NET.Native.Framework.1.3       
#Microsoft.NET.Native.Framework.1.3       
#Microsoft.NET.Native.Runtime.1.3         
##Microsoft.NET.Native.Runtime.1.3         
##Microsoft.NET.Native.Runtime.1.4         
##Microsoft.NET.Native.Runtime.1.4         
"Microsoft.OneConnect",                     
"Microsoft.Services.Store.Engagement",      
"Microsoft.StorePurchaseApp",
#Microsoft.VCLibs.140.00                  
#Microsoft.VCLibs.140.00                  
#Microsoft.WindowsAlarms
#Microsoft.WindowsCalculator              
"Microsoft.WindowsFeedbackHub",
"Microsoft.WindowsMaps",                    
"Microsoft.WindowsStore",
"Microsoft.XboxIdentityProvider"
#"Microsoft.XboxIdentityProvider"
#Microsoft.Getstarted                     
#Microsoft.WindowsSoundRecorder           
#Microsoft.MicrosoftOfficeHub


ForEach ($App in $AppsList) 
{ 
    $PackageFullName = (Get-AppxPackage -AllUsers $App).PackageFullName
    $ProPackageFullName = (Get-AppxProvisionedPackage -online | where {$_.Displayname –like "*$app*"}).PackageName

 

    if ($ProPackageFullName) 
    { 
        Remove-AppxProvisionedPackage -online -packagename $ProPackageFullName 
    } 

}