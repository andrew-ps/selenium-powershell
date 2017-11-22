[System.Reflection.Assembly]::LoadFrom("$PSScriptRoot\WebDriver.dll")
[System.Reflection.Assembly]::LoadFrom("$PSScriptRoot\WebDriver.Support.dll")
[System.Reflection.Assembly]::LoadFrom("$PSScriptRoot\Selenium.WebDriverBackedSelenium.dll")

function Start-SEChrome {
    New-Object -TypeName "OpenQA.Selenium.Chrome.ChromeDriver"
}

function Start-SEFirefox {
    New-Object -TypeName "OpenQA.Selenium.Firefox.FirefoxDriver"
}

function Stop-SEDriver {
    param($Driver)

    $Driver.Dispose()
}

function Enter-SEUrl {
    param($Driver, $Url)

    $Driver.Navigate().GoToUrl($Url)
}

function Find-SEElement {
    param(
        [Parameter()]
        $Driver,
        [Parameter()]
        $Element,
        [Parameter(ParameterSetName = "ByName")]
        $Name,
        [Parameter(ParameterSetName = "ById")]
        $Id,
        [Parameter(ParameterSetName = "ByClassName")]
        $ClassName,
        [Parameter(ParameterSetName = "ByLinkText")]
        $LinkText,
        [Parameter(ParameterSetName = "ByTagName")]
        $TagName)

    Process {

        if ($Driver -ne $null -and $Element -ne $null) {
            throw "Driver and Element may not be specified together."
        }
        elseif ($Driver -ne $Null) {
            $Target = $Driver
        }
        elseif ($Element -ne $Null) {
            $Target = $Element
        }
        else {
            "Driver or element must be specified"
        }

        if ($PSCmdlet.ParameterSetName -eq "ByName") {
            $Target.FindElements([OpenQA.Selenium.By]::Name($Name))
        }

        if ($PSCmdlet.ParameterSetName -eq "ById") {
            $Target.FindElements([OpenQA.Selenium.By]::Id($Id))
        }

        if ($PSCmdlet.ParameterSetName -eq "ByLinkText") {
            $Target.FindElements([OpenQA.Selenium.By]::LinkText($LinkText))
        }

        if ($PSCmdlet.ParameterSetName -eq "ByClassName") {
            $Target.FindElements([OpenQA.Selenium.By]::ClassName($ClassName))
        }

        if ($PSCmdlet.ParameterSetName -eq "ByTagName") {
            $Target.FindElements([OpenQA.Selenium.By]::TagName($TagName))
        }
    }
}

function Invoke-SEClick {
    param([OpenQA.Selenium.IWebElement]$Element)

    $Element.Click()
}

function Send-SEKeys {
    param([OpenQA.Selenium.IWebElement]$Element, [string]$Keys)

    $Element.SendKeys($Keys)
}

function Get-SECookie {
    param($Driver)

    $Driver.Manage().Cookies.AllCookies.GetEnumerator()
}

function Remove-SECookie {
    param($Driver)

    $Driver.Manage().Cookies.DeleteAllCookies()
}

function Set-SECookie {
    param($Driver, $name, $value)

    $cookie = New-Object -TypeName OpenQA.Selenium.Cookie -ArgumentList $Name,$value

    $Driver.Manage().Cookies.AddCookie($cookie)
}

function Get-SEElementAttribute {
    param(
        [Parameter(ValueFromPipeline=$true, Mandatory = $true)]
        [OpenQA.Selenium.IWebElement]$Element,
        [Parameter(Mandatory=$true)]
        [string]$Attribute
    )

    Process {
        $Element.GetAttribute($Attribute)
    }
}

<#
 .Synopsis
  Gets Multiple Selenium Web Elements.

 .Description
  This command will allow you to find multiple elements in your Selenium WebDriver
  object. It will also return a single element.

 .Parameter Driver
  Selenium WebDriver Object.

 .Parameter Element
  Selenium Element Object

 .Parameter FindBy
  Specify the Method to Find the Selenium Element
  Availble Types are : ClassName, CssSelector, Id, LinkText, Name, PartialLinkText, TagName, XPath
  You will need to determine which method to use by what is unique to the element you are looking
  for by inspecting it's HTML.

 .Parameter Value
  Specify the Value for the FindBy Parameter. The Value is Case Sensative to what is in HTML.

 .Example
  # Display a List of SubReddits Currently Visible on the Top Bar
  $subRedditBar = Find-SEElement -Driver $driver -FindBy ClassName -Value 'sr-list'
  $subReddits = Find-SEElements -Element $subRedditBar -FindBy ClassName -Value 'choice'
  $subReddits | where {$_.Displayed -eq $true} | select -ExpandProperty text

 .Example
  # Get a Table Element and then Get Each Row in a Table
  $tables = Find-SEElement -Driver $driver -FindBy TagName -Value 'table'
  $tableRows = Find-SEElements -Element $tables -FindBy TagName -Value 'tr'
  $tableRows

  .Notes
#>
function Find-SEElements {
    [CmdletBinding(DefaultParameterSetName='Driver')]
    param(
    [Parameter(Mandatory=$true, ParameterSetName = 'Driver')]
    [OpenQA.Selenium.Remote.RemoteWebDriver]$Driver,
    [Parameter(Mandatory=$true, ParameterSetName = 'Element')]
    [OpenQA.Selenium.Remote.RemoteWebElement]$Element,
    [ValidateSet("ClassName","CssSelector","Id","LinkText","Name","PartialLinkText","TagName","XPath")]
    [Parameter(Mandatory=$true)]
    [string]$FindBy,
    [Parameter(Mandatory=$true)]
    [string]$Value
    )
    if($Driver){
        try{
            $webElements = $Driver."FindElementsBy$FindBy"($Value)
        }
        catch{
            return $null
        }
    }
    else{
        try{
            $webElements = $Element."FindElementsBy$FindBy"($Value)
        }
        catch{
            return $null
        }
    }
    return $webElements
}
<#
 .Synopsis
  Gets a Selenium Web Element.

 .Description
  This command will allow you to find an element in your Selenium WebDriver
  object so you can click, send keys, or check attributes on the element.
  This returns a null value if no ELement is found.

 .Parameter Driver
  Selenium WebDriver Object.

 .Parameter Element
  Selenium Element Object

 .Parameter FindBy
  Specify the Method to Find the Selenium Element
  Availble Types are : ClassName, CssSelector, Id, LinkText, Name, PartialLinkText, TagName, XPath
  You will need to determine which method to use by what is unique to the element you are looking for by inspecting it's HTML

 .Parameter Value
  Specify the Value for the FindBy Parameter. The Value is Case Sensative to what is in HTML.

 .Example
  # Get the Reddit.com Search Bar Element and Input a String
  # HTML on Search Bar : <input type="text" name="q" placeholder="search" tabindex="20">
  $searchBar = Find-SEElement -Driver $driver -FindBy Name -Value 'q'
  $searchBar.SendKeys('PowerCLI Module')

 .Example
  # Get a Table Element and then Display the Second Row
  $table = Find-SEElement -Driver $driver -FindBy TagName -Value 'table'
  $tableRows = Find-SEElements -Element $table -FindBy TagName -Value 'tr'
  $tableRows[1]

  .Notes
#>
function Find-SEElement {
    [CmdletBinding(DefaultParameterSetName='Driver')]
    param(
    [Parameter(Mandatory=$true, ParameterSetName = 'Driver')]
    [OpenQA.Selenium.Remote.RemoteWebDriver]$Driver,
    [Parameter(Mandatory=$true, ParameterSetName = 'Element')]
    [OpenQA.Selenium.Remote.RemoteWebElement]$Element,
    [ValidateSet("ClassName","CssSelector","Id","LinkText","Name","PartialLinkText","TagName","XPath")]
    [Parameter(Mandatory=$true)]
    [string]$FindBy,
    [Parameter(Mandatory=$true)]
    [string]$Value
    )
    if($Driver){
        try{
            $webElement = $Driver."FindElementsBy$FindBy"($Value)
            if($webElement.count -eq 1){
                $webElement = $Driver."FindElementBy$FindBy"($Value)
            }
            else{
                throw "$($webElement.count) Elements Found Try to Narrow Search or Use Get-SEElements Cmdlet"
            }
        }
        catch{
            return $null
        }
    }
    else{
        try{
            $webElement = $Element."FindElementsBy$FindBy"($Value)
            if($webElement.count -eq 1){
                $webElement = $Element."FindElementBy$FindBy"($Value)
            }
            else{
                throw "$($webElement.count) Elements Found Try to Narrow Search or Use Get-SEElements Cmdlet"
            }
        }
        catch{
            return $null
        }
    }
    return $webElement
}
<#
 .Synopsis
  Clicks a Selenium WebElement

 .Description
  This command will select or click on a Selenium Element

 .Parameter Driver
  Selenium WebDriver Object.

 .Parameter Element
  Selenium Element Object

 .Parameter FindBy
  Specify the Method to Find the Selenium Element
  Availble Types are : ClassName, CssSelector, Id, LinkText, Name, PartialLinkText, TagName, XPath
  You will need to determine which method to use by what is unique to the element you are looking for by inspecting it's HTML
  This site may help you figure out inspecting elements a bit : https://www.lifewire.com/get-inspect-element-tool-for-browser-756549

 .Parameter Value
  Specify the Value for the FindBy Parameter. This is Case Sensative.

 .Example
  # Switch to "New" Tab on Reddit.com
  Select-SEElement -Driver $driver -FindBy LinkText -Value 'new'

 .Example
  # Login to Reddit.com
  Write-SEElement -Driver $driver -FindBy Name -Value 'user' -Text 'USERNAME'
  Write-SEElement -Driver $driver -FindBy Name -Value 'passwd' -Text 'PASSWORD'
  Select-SEElement -Driver $driver -FindBy ClassName -Value 'btn'

  .Notes
#>
function Select-SEElement {
    [CmdletBinding(DefaultParameterSetName='Driver')]
    param(
    [Parameter(Mandatory=$true, ParameterSetName = 'Driver')]
        [OpenQA.Selenium.Remote.RemoteWebDriver]$Driver,
    [Parameter(Mandatory=$true, ParameterSetName = 'Element')]
    [Parameter(Mandatory=$true, ParameterSetName = 'Pipeline',ValueFromPipeline=$true)]
        [OpenQA.Selenium.Remote.RemoteWebElement]$Element,
    [ValidateSet("ClassName","CssSelector","Id","LinkText","Name","PartialLinkText","TagName","XPath")]
    [Parameter(Mandatory=$true, ParameterSetName = 'Driver')]
        [string]$FindBy,
    [Parameter(Mandatory=$true, ParameterSetName = 'Driver')]
        [string]$Value
    )
    if($Driver){
        try{
            $webElement = $Driver."FindElementsBy$FindBy"($Value)
            if($webElement.count -eq 1){
                $webElement = $Driver."FindElementBy$FindBy"($Value)
            }
            else{
                throw 'Multiple elements found'
            }
        }
        catch{
            throw 'Element not found'
        }
    }
    else{
        $webElement = $Element
    }
    $webElement.Click()
}
<#
 .Synopsis
  Sends Input to a Selenium Element.

 .Description
  This command will allow you to send input to any form or textbox.

 .Parameter Driver
  Selenium WebDriver Object.

 .Parameter Element
  Selenium Element Object

 .Parameter FindBy
  Specify the Method to Find the Selenium Element
  Availble Types are : ClassName, CssSelector, Id, LinkText, Name, PartialLinkText, TagName, XPath
  You will need to determine which method to use by what is unique to the element you are looking for by inspecting it's HTML
  This site may help you figure out inspecting elements a bit : https://www.lifewire.com/get-inspect-element-tool-for-browser-756549

 .Parameter Value
  Specify the Value for the FindBy Parameter. This is Case Sensative.

 .Parameter Text
  Specify the Value to be Inputted into the Element.

 .Example
  # Get the Reddit.com Search Bar Element and Input a String
  # HTML : <input type="text" name="q" placeholder="search" tabindex="20">
  Get-SEElement -Driver $driver -FindBy Name -Value 'q' | Write-SEElement -Text 'PowerCLI Module'

 .Example
  # Login to Reddit.com
  Write-SEElement -Driver $driver -FindBy Name -Value 'user' -Text 'USERNAME'
  Write-SEElement -Driver $driver -FindBy Name -Value 'passwd' -Text 'PASSWORD'
  Select-SEElement -Driver $driver -FindBy ClassName -Value 'btn'

  .Notes
#>
function Write-SEElement {
    [CmdletBinding(DefaultParameterSetName='Driver')]
    param(
    [Parameter(Mandatory=$true, ParameterSetName = 'Driver')]
    [OpenQA.Selenium.Remote.RemoteWebDriver]$Driver,
    [Parameter(Mandatory=$true, ParameterSetName = 'Element')]
    [Parameter(Mandatory=$true, ParameterSetName = 'Pipeline',ValueFromPipeline=$true)]
    [OpenQA.Selenium.Remote.RemoteWebElement]$Element,
    [ValidateSet("ClassName","CssSelector","Id","LinkText","Name","PartialLinkText","TagName","XPath")]
    [Parameter(Mandatory=$true, ParameterSetName = 'Driver')]
    [string]$FindBy,
    [Parameter(Mandatory=$true, ParameterSetName = 'Driver')]
    [string]$Value,
    [Parameter(Mandatory=$true, ParameterSetName = 'Driver')]
    [Parameter(Mandatory=$true, ParameterSetName = 'Pipeline',Position=0)]
    [string]$Text
    )
    if($Driver){
        try{
            $webElement = $Driver."FindElementsBy$FindBy"($Value)
            if($webElement.count -eq 1){
                $webElement = $Driver."FindElementBy$FindBy"($Value)
            }
        }
        catch{
            return $false
        }
    }
    else{
        $webElement = $Element
    }
    $webElement.SendKeys("$Text")
}