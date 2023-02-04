#Function to copy attachments between list items
#See: https://www.mrsharepoint.guru/sharepoint-copy-list-item-to-another-site/
#
Function Copy-SPOAttachments()
{
    param
    (
        [Parameter(Mandatory=$true)] [Microsoft.SharePoint.Client.ListItem] $SourceItem,
        [Parameter(Mandatory=$true)] [Microsoft.SharePoint.Client.ListItem] $DestinationItem
    )
    Try {
        #Get All Attachments from Source list items
        $Attachments = Get-PnPProperty -ClientObject $SourceItem -Property "AttachmentFiles" -Connection $SourceConn
        $Attachments | ForEach-Object {
            #Download the Attachment to Temp
            $File  = Get-PnPFile -Connection $SourceConn -Url $_.ServerRelativeUrl -FileName $_.FileName -Path $Env:TEMP -AsFile -Force
            #Add Attachment to Destination List Item
            $FileStream = New-Object IO.FileStream(($Env:TEMP+"\"+$_.FileName),[System.IO.FileMode]::Open)  
            $AttachmentInfo = New-Object -TypeName Microsoft.SharePoint.Client.AttachmentCreationInformation
            $AttachmentInfo.FileName = $_.FileName
            $AttachmentInfo.ContentStream = $FileStream
            $AttachFile = $DestinationItem.AttachmentFiles.Add($AttachmentInfo)
            Invoke-PnPQuery -Connection $DestinationConn
       
            #Delete the Temporary File
            Remove-Item -Path $Env:TEMP\$($_.FileName) -Force
        }
    }
    Catch {
        write-host -f Red "Error Copying Attachments:" $_.Exception.Message
    }
}
 
#Function to copy list items from one list to another
Function Copy-SPOListItems()
{
    param
    (
        [Parameter(Mandatory=$true)] [Microsoft.SharePoint.Client.List] $SourceList,
        [Parameter(Mandatory=$true)] [Microsoft.SharePoint.Client.List] $DestinationList
    )
    Try {
        #Get All Items from the Source List in batches 
        Write-Progress -Activity "Reading Source..." -Status "Getting Items from Source List. Please wait..."
        $SourceListItems = Get-PnPListItem -List $SourceList -PageSize 500 -Connection $SourceConn
        $SourceListItemsCount= $SourceListItems.count
        Write-host "Total Number of Items Found:"$SourceListItemsCount     
   
        #Get fields to Update from the Source List - Skip Read only, hidden fields, content type and attachments
        $SourceListFields = Get-PnPField -List $SourceList -Connection $SourceConn | Where { (-Not ($_.ReadOnlyField)) -and (-Not ($_.Hidden)) -and ($_.InternalName -ne  "ContentType") -and ($_.InternalName -ne  "Attachments") }
 
        #Loop through each item in the source and Get column values, add them to Destination
        [int]$Counter = 1
        ForEach($SourceItem in $SourceListItems)
        {  
            $ItemValue = @{}
            #Map each field from source list to Destination list
            Foreach($SourceField in $SourceListFields)
            {
                #Check if the Field value is not Null
                If($SourceItem[$SourceField.InternalName] -ne $Null)
                {
                    #Handle Special Fields
                    $FieldType  = $SourceField.TypeAsString                    
   
                    If($FieldType -eq "User" -or $FieldType -eq "UserMulti") #People Picker Field
                    {
                        $PeoplePickerValues = $SourceItem[$SourceField.InternalName] | ForEach-Object { $_.Email}
                        $ItemValue.add($SourceField.InternalName,$PeoplePickerValues)
                    }
                    ElseIf($FieldType -eq "Lookup" -or $FieldType -eq "LookupMulti") # Lookup Field
                    {
                        $LookupIDs = $SourceItem[$SourceField.InternalName] | ForEach-Object { $_.LookupID.ToString()}
                        $ItemValue.add($SourceField.InternalName,$LookupIDs)
                    }
                    ElseIf($FieldType -eq "URL") #Hyperlink
                    {
                        $URL = $SourceItem[$SourceField.InternalName].URL
                        $Description  = $SourceItem[$SourceField.InternalName].Description
                        $ItemValue.add($SourceField.InternalName,"$URL, $Description")
                    }
                    ElseIf($FieldType -eq "TaxonomyFieldType" -or $FieldType -eq "TaxonomyFieldTypeMulti") #MMS
                    {
                        $TermGUIDs = $SourceItem[$SourceField.InternalName] | ForEach-Object { $_.TermGuid.ToString()}                    
                        $ItemValue.add($SourceField.InternalName,$TermGUIDs)
                    }
                    Else
                    {
                        #Get Source Field Value and add to Hashtable                        
                        $ItemValue.add($SourceField.InternalName,$SourceItem[$SourceField.InternalName])
                    }
                }
            }
            #Copy Created by, Modified by, Created, Modified Metadata values
            $ItemValue.add("Created", $SourceItem["Created"]);
            $ItemValue.add("Modified", $SourceItem["Modified"]);
            $ItemValue.add("Author", $SourceItem["Author"].Email);
            $ItemValue.add("Editor", $SourceItem["Editor"].Email);
 
            Write-Progress -Activity "Copying List Items:" -Status "Copying Item ID '$($SourceItem.Id)' from Source List ($($Counter) of $($SourceListItemsCount))" -PercentComplete (($Counter / $SourceListItemsCount) * 100)
             
            #Copy column value from Source to Destination
            $NewItem = Add-PnPListItem -List $DestinationList -Values $ItemValue -Connection $DestinationConn
   
            #Copy Attachments
            Copy-SPOAttachments -SourceItem $SourceItem -DestinationItem $NewItem
   
            Write-Host "Copied Item ID from Source to Destination List:$($SourceItem.Id) ($($Counter) of $($SourceListItemsCount))"
            $Counter++
        }
   }
    Catch {
        Write-host -f Red "Error:" $_.Exception.Message 
    }
}
   
#Set Parameters
$SourceSiteURL = "https://iptrack.sharepoint.com/sites/RESTAPI"
$SourceListName = "Recycling-Duties"
  
$DestinationSiteURL = "https://iptrack.sharepoint.com/sites/RESTAPI/public"
$DestinationListName = "Duties"
  
#Connect to Source and destination sites
$SourceConn = Connect-PnPOnline -Url $SourceSiteURL -Interactive -ReturnConnection
$SourceList = Get-PnPList -Identity $SourceListName -Connection $SourceConn
  
$DestinationConn = Connect-PnPOnline -Url $DestinationSiteURL -Interactive -ReturnConnection
$DestinationList = Get-PnPList -Identity $DestinationListName -Connection $DestinationConn
   
#Call the Function to Copy List Items between Lists
Copy-SPOListItems -SourceList $SourceList -DestinationList $DestinationList