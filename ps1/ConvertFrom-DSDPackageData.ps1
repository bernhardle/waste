#
param([String] $psmodules = "C:\Users\Bernhard\Eigene Dateien\XML Formulare\Common\ps1")
#
function local:convertDSDPackageData([String] $inp = $(throw "Convert-DSDPackageData.ps1::convertDSDPackageData(...) Missing input file path."), [String] $oup = $(throw "Convert-DSDPackageData.ps1::convertDSDPackageData(...) Missing output file path.")) {
	#
	[Reflection.Assembly] $local:ass = [Reflection.Assembly]::LoadWithPartialName("System.Xml")
	#
	[String[]] $header=@("ean","material","assigned","validfrom","validto","description","contract","glass","ppk","tin","alloy","plastics","laminates","aggregates","natural")
	#
	[Xml] $xml = $(ConvertTo-XML -NoTypeInformation -As "Document" -InputObject (Import-Csv -Header $header -Encoding 'oem' -Delimiter ';' $inp) -Depth 6)
	#
	[System.Xml.XmlWriterSettings] $local:xws = New-Object System.Xml.XmlWriterSettings
	#
	[System.Xml.XmlWriter] $local:xw = [System.Xml.XmlWriter]::Create($oup, $xws)
	#
	$xw.WriteStartElement("root","http://www.rothenberger.com/dsd") ; 
	#
	$xw.WriteRaw($xml.Objects.Object.InnerXml) 
	#
	$xw.WriteEndElement() 
	#
	$xw.Close()
	#
	$xw.Dispose()
	#
}
[String] $local:InputPath = $(. OpenFileDialog -title 'Select input file for reading DSD package data' -type 'csv' -defpath $pwd)
#
if ([System.IO.File]::Exists($local:InputPath) -eq $true) {
	#
	[String] $local:OutputPath = $(. $psmodules\SaveFileDialog.ps1 -title 'Select output file for writing XML converted DSD package data' -defpath $pwd -delete $true -type 'xml')
	#
	. convertDSDPackageData -inp $local:InputPath -oup $local:OutputPath
	#
}
Write-Host -Message @"
Bye bye ...
"@
