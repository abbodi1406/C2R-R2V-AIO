# Office Click-to-Run Retail-to-Volume AIO

- A standalone fork of C2R-R2V,  
where cleanospp.exe is embedded in the script and get extracted upon execution using powershell,  
which makes it easier to move and distribute the script alone

- Convert already-installed Office ClickToRun licensing from Retail to Volume

- This is not an activator, just a licensing converter

- Supports: Microsoft 365, Office 2024, Office 2021, Office 2019, Office 2016, Office 2013

- Activated Retail products will be skipped from conversion  
this include valid Office 365 subscriptions, or perpetual Office (MAK, OEM, MSDN, Retail..)

- Current Office licenses will be cleaned up (unless retail-activated Office detected)  
then, proper Volume licenses will be installed based on the detected Product IDs

- Office Mondo suite cover all products, if detected, only its licenses will be installed

- Microsoft Office 365 products will be converted with Mondo licenses by default  
also, corresponding Office 365 Retail Grace Key will be installed

- Windows 10/11: Office 2016 products will be converted with corresponding Office 2019 licenses (if RTM detected)

- Windows 8.1: Office 2016/2019 products will be converted with corresponding Office 2021 licenses (if RTM detected)

- Office Professional suite will be converted with Office Professional Plus licenses

- Office HomeBusiness/HomeStudent suites will be converted with Office Standard licenses

- Office 2013 products follow the same logic, but handled separately

- If main products Suites are detected, single apps licenses will not be installed to avoid duplication

- Suites:  
O365ProPlus, O365Business, O365SmallBusPrem, O365HomePrem, O365EduCloud  
ProPlus, Professional, Standard, HomeBusiness, HomeStudent, Visio, Project

- Apps:  
Access, Excel, InfoPath, Onenote, Outlook, PowerPoint, Publisher, SkypeForBusiness, Word, Groove (OneDrive for Business)

- Suites <> Apps inclusion:  
O365ProPlus, O365Business, O365SmallBusPrem, ProPlus: all apps  
O365HomePrem, Professional: all apps except SkypeForBusiness  
Standard: all apps except Access, SkypeForBusiness  
HomeBusiness: Excel, OneNote, PowerPoint, Word, Outlook  
HomeStudent, O365EduCloud: Excel, OneNote, PowerPoint, Word

## Remarks

- On Windows 7, Office C2R 16.0 licensing service require Universal C Runtime to work correctly

- UCRT is available in the latest Monthly Rollup, or the separate update KB3118401

- Additionally, Office programs themselves require recent Windows 7 updates to start properly

- While Office 2019/2021 can be installed on Windows 7/8.1, it will not function or work properly

## Office 2019/2021

- Office 2019/2021 products are officially blocked on Windows 7 and 8.1  
to workaround that, use the following steps

- If you want Office Professional Plus:  
1) install Office 2016 SKU ProPlusRetail  
2) run C2R-R2V-AIO.cmd

- Additionally, if you want Project and/or Visio:  
1) install desired Office 2016 SKU: ProjectProRetail, ProjectStdRetail, VisioProRetail, VisioStdRetail  
2) run C2R-R2V-AIO.cmd

## Manual advanced options:

- To run the process in debug mode "with" conversion  
edit C2R-R2V-AIO.cmd and set _Debug=1

- To run the process in debug mode "without" conversion  
edit C2R-R2V-AIO.cmd and set _Cnvrt=0

## Credits

- @Windows_Addict / Features suggestion, testing and co-enhancing  
- @AveYo          / [Compressed2TXT](https://github.com/AveYo/Compressed2TXT)  
- @ratzlefatz     / [OfficeRTool](https://forums.mydigitallife.net/posts/1125229/)
