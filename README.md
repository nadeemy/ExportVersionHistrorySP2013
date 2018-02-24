# ExportVersionHistrorySP2013

Export version history can be used to export the version history of SharePoint 201 List items to Microsoft Excel. It provides button in Ribbon to do bulk export.

# Deployment steps

**Add Solution using stsadm.exe:**
stsadm.exe -o addsolution -filename C:\NY.ExportVersionHistory.wsp 

**Deploy Solution using stsadm.exe:**
stsadm.exe -o deploysolution -name NY.ExportVersionHistory.wsp -immediate -allowgacdeployment

**Add Solution using PowerShell:**
Add-SPSolution -LiteralPath c:\NY.ExportVersionHistory.wsp

**Deploy Solution using PowerShell:**
Install-SPSolution -Identity NY.ExportVersionHistory.wsp -GACDeployment

## Feature Activation

Go to the site collection where you want to activate the feature. The feature is scoped at Site collection level.

For more information please visit http://www.sharepointnadeem.com/2012/07/export-version-history-of-sharepoint.html

## WANT TO SHOW APPRECIATION? 

If you find this tool useful and want to show your appreciation, you can go to my blog [SharePoint Learnings](http://www.sharepointnadeem.com/2012/07/export-version-history-of-sharepoint.html) and click the banners on the blog to visit my blog sponsors or you could spread the word about it on social media sites.

