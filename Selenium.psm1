[System.Reflection.Assembly]::LoadFrom("$PSScriptRoot\WebDriver.dll")
[System.Reflection.Assembly]::LoadFrom("$PSScriptRoot\WebDriver.Support.dll")
[System.Reflection.Assembly]::LoadFrom("$PSScriptRoot\Selenium.WebDriverBackedSelenium.dll")

<#
 .Synopsis
  Starts a New Chrome Driver.

 .Description
  This command will start a new Chrome Driver. You can also specify Chrome options,
  and a starting URL if desired.

 .Parameter URL
  URL to navigate to. Must have a http://, https://, ftp:// etc.

 .Parameter Options
  Specify Chrome Driver Options. List of options available here :
  https://peter.sh/experiments/chromium-command-line-switches/

 .Example
  # Create new driver with Chrome Options.
  $driverOptions = New-Object OpenQA.Selenium.Chrome.ChromeOptions
  # Launches Driver in Full Screen Kiosk Mode, and disables 'Chrome is being controlled by automated test software'
  $driverOptions.AddArguments("start-fullscreen", "kiosk", "disable-infobars")
  # This should keep the browser running after the driver closes but does not seem to be working at this time
  $driverOptions.LeaveBrowserRunning = $true
  # This would disable the prompt to save credentials on login forms
  $driverOptions.AddUserProfilePreference("credentials_enable_service", $false)
  $driver = Start-SEChrome -Url 'http://google.com' -Options $driverOptions

  .Notes
#>
function Start-SEChrome {
    param(
        [Parameter(Mandatory=$false)]
        $Url,
        [Parameter(Mandatory=$false)]
        [OpenQA.Selenium.Chrome.ChromeOptions]$Options
    )
    if($options -ne $null){
        $driver = New-Object OpenQA.Selenium.Chrome.ChromeDriver($Options)
    }
    else{
        $driver = New-Object -TypeName "OpenQA.Selenium.Chrome.ChromeDriver"
    }
    if($Url){
        $driver.Navigate().GoToUrl($Url)
    }
    return $driver
}

<#
 .Synopsis
  Starts a New Firefox Driver.

 .Description
  This command will start a new Firefox Driver.

 .Example
  # Starts new Firefox Driver
  $driver = Start-SEFirefox

  .Notes
#>
function Start-SEFirefox {
    New-Object -TypeName "OpenQA.Selenium.Firefox.FirefoxDriver"
}

<#
 .Synopsis
  Stops a Selenium Driver.

 .Description
  This command will dispose of the Selenium Driver and close out the browser.
  # You can keep the browser open by specifying ####

 .Parameter Driver
  Selenium WebDriver Object.

 .Example
  # Open a new driver and close it back out
  $driver = Start-SEChrome -Url 'http://google.com' -Options $driverOptions
  Stop-SEDriver -Driver $driver

  .Notes
#>
function Stop-SEDriver {
    param($Driver)

    $Driver.Dispose()
}

<#
 .Synopsis
  Goes to a specified URL.

 .Description
  This command will go to the URL specified. You need to put in http://, https://,
  ftp:// etc.

 .Parameter Driver
  Selenium WebDriver Object.

 .Parameter URL
  URL to navigate to. Must have a http://, https://, ftp:// etc. address.

 .Example
  # Create new driver and navigate to a website.
  $driver = Start-SEChrome
  Enter-SEUrl -Driver $driver -Url 'http://google.com'

  .Notes
#>
function Enter-SEUrl {
    param(
        [Parameter(Position=0,Mandatory=$true)]
        [OpenQA.Selenium.Remote.RemoteWebDriver]$Driver,
        [Parameter(Mandatory=$true)]
        $Url
    )

    $Driver.Navigate().GoToUrl($Url)
}

<#
 .Synopsis
  Gets a Selenium Web Element.

 .Description
  This command will allow you to find an element in your Selenium WebDriver
  object so you can click, send keys, or check attributes on the element.
  You can also specify a Selenium Element such as a division or table to
  limit your search to a sub-element within that tag.
  This returns an error if no element is found.

 .Parameter Driver
  Selenium WebDriver Object.

 .Parameter Element
  Selenium Element Object.

 .Parameter Name
  Specify an element by name. This is case sensative.

 .Parameter Id
  Specify an element by id. This is case sensative.

 .Parameter ClassName
  Specify an element by class name. This is case sensative.

 .Parameter LinkText
  Specify an element by link text.

 .Parameter TagName
  Specify an element by tag name.

 .Parameter PartialLinkText
  Specify an element by partial link text.

 .Parameter CssSelector
  Specify an element by css selector. This is case sensative.

 .Parameter XPath
  Specify an element by xpath.

 .Example
  # Get the Reddit.com Search Bar Element and Input a String
  # HTML on Search Bar : <input type="text" name="q" placeholder="search" tabindex="20">
  $searchBar = Find-SEElement -Driver $driver -Name 'q'
  $searchBar.SendKeys('PowerCLI Module')

 .Example
  # Get a Table Element and then Display the attributes of the Second Row
  $table = Find-SEElement -Driver $driver -TagName 'table'
  $tableRows = Find-SEElements -Element $table -TagName 'tr'
  $tableRows[1]

  .Notes
#>
function Find-SEElement {
    param(
        [Parameter(Position=0, ParameterSetName = "DriverByName")]
        [Parameter(Position=0, ParameterSetName = "DriverById")]
        [Parameter(Position=0, ParameterSetName = "DriverByClassName")]
        [Parameter(Position=0, ParameterSetName = "DriverByLinkText")]
        [Parameter(Position=0, ParameterSetName = "DriverByTagName")]
        [Parameter(Position=0, ParameterSetName = "DriverByPartialLinkText")]
        [Parameter(Position=0, ParameterSetName = "DriverByCssSelector")]
        [Parameter(Position=0, ParameterSetName = "DriverByXPath")]
        [OpenQA.Selenium.Remote.RemoteWebDriver]$Driver,
        [Parameter(Position=0, ParameterSetName = "ElementByName")]
        [Parameter(Position=0, ParameterSetName = "ElementById")]
        [Parameter(Position=0, ParameterSetName = "ElementByClassName")]
        [Parameter(Position=0, ParameterSetName = "ElementByLinkText")]
        [Parameter(Position=0, ParameterSetName = "ElementByTagName")]
        [Parameter(Position=0, ParameterSetName = "ElementByPartialLinkText")]
        [Parameter(Position=0, ParameterSetName = "ElementByCssSelector")]
        [Parameter(Position=0, ParameterSetName = "ElementByXPath")]
        [OpenQA.Selenium.Remote.RemoteWebElement]$Element,
        [Parameter(ParameterSetName = "DriverByName")]
        [Parameter(ParameterSetName = "ElementByName")]
        $Name,
        [Parameter(ParameterSetName = "DriverById")]
        [Parameter(ParameterSetName = "ElementById")]
        $Id,
        [Parameter(ParameterSetName = "DriverByClassName")]
        [Parameter(ParameterSetName = "ElementByClassName")]
        $ClassName,
        [Parameter(ParameterSetName = "DriverByLinkText")]
        [Parameter(ParameterSetName = "ElementByLinkText")]
        $LinkText,
        [Parameter(ParameterSetName = "DriverByTagName")]
        [Parameter(ParameterSetName = "ElementByTagName")]
        $TagName,
        [Parameter(ParameterSetName = "DriverByPartialLinkText")]
        [Parameter(ParameterSetName = "ElementByPartialLinkText")]
        $PartialLinkText,
        [Parameter(ParameterSetName = "DriverByCssSelector")]
        [Parameter(ParameterSetName = "ElementByCssSelector")]
        $CssSelector,
        [Parameter(ParameterSetName = "DriverByXPath")]
        [Parameter(ParameterSetName = "ElementByXPath")]
        $XPath
        )

    Process {

        if ($Driver -ne $Null) {
            $Target = $Driver
        }
        elseif ($Element -ne $Null) {
            $Target = $Element
        }
        else {
            throw "Driver or element must be specified"
        }

        if ($PSCmdlet.ParameterSetName -match "ByName") {
            try{
                $Target.FindElements([OpenQA.Selenium.By]::Name($Name))
            }
            catch{
                throw "No Element Found"
            }
        }

        if ($PSCmdlet.ParameterSetName -match "ById") {
            try{
                $Target.FindElements([OpenQA.Selenium.By]::Id($Id))
            }
            catch{
                throw "No Element Found"
            }
        }

        if ($PSCmdlet.ParameterSetName -match "ByLinkText") {
            try{
                $webElement = $Target.FindElements([OpenQA.Selenium.By]::LinkText($LinkText))
            }
            catch{
                throw "No Element Found"
            }
            if($webElement.Count -eq 0){
                throw "No Element Found"
            }
            else{
                return $webElement
            }
        }

        if ($PSCmdlet.ParameterSetName -match "ByClassName") {
            try{
                $Target.FindElements([OpenQA.Selenium.By]::ClassName($ClassName))
            }
            catch{
                throw "No Element Found"
            }
        }

        if ($PSCmdlet.ParameterSetName -match "ByTagName") {
            try{
                $Target.FindElements([OpenQA.Selenium.By]::TagName($TagName))
            }
            catch{
                throw "No Element Found"
            }
        }

        if ($PSCmdlet.ParameterSetName -match "ByPartialLinkText") {
            try{
                $Target.FindElements([OpenQA.Selenium.By]::PartialLinkText($TagName))
            }
            catch{
                throw "No Element Found"
            }
        }

        if ($PSCmdlet.ParameterSetName -match "ByCssSelector") {
            try{
                $Target.FindElements([OpenQA.Selenium.By]::CssSelector($TagName))
            }
            catch{
                throw "No Element Found"
            }
        }

        if ($PSCmdlet.ParameterSetName -match "ByXPath") {
            try{
                $Target.FindElements([OpenQA.Selenium.By]::XPath($TagName))
            }
            catch{
                throw "No Element Found"
            }
        }
    }
}

<#
 .Synopsis
  Clicks a Selenium WebElement

 .Description
  This command will click on a Selenium Element

 .Parameter Element
  Selenium Element Object

 .Example
  # Login to Reddit.com
  $username = Find-SEElement -Name 'user'
  $password = Find-SEElement -Name 'passwd'
  $submitBtn = Find-SEElement -ClassName 'btn'
  Send-SEKeys -Element $username -Keys 'USERNAME'
  Send-SEKeys -Element $password -Keys 'PASSWORD'
  Invoke-SEClick -Element $submitBtn

  .Example
  # Login to Reddit.com using pipeline
  Find-SEElement -Name 'user' | Send-SEKeys -Keys 'USERNAME'
  Find-SEElement -Name 'passwd' | Send-SEKeys -Keys 'PASSWORD'
  Find-SEElement -ClassName 'btn' | Invoke-SEClick

  .Notes
#>
function Invoke-SEClick {
    param(
    [Parameter(Position=0,Mandatory=$true,ValueFromPipeline=$true)]
    [OpenQA.Selenium.IWebElement]$Element
    )

    $Element.Click()
}

<#
 .Synopsis
  Sends keys to a Selenium WebElement

 .Description
  This command will send keys/input to a Selenium Element.

 .Parameter Element
  Selenium Element Object

 .Parameter Keys
  Specify the Keys to Send to the Element.

 .Parameter SubmitForm
  Attempts to Submit Form if its a Form Element.

 .Example
  # Login to Reddit.com
  $username = Find-SEElement -Name 'user'
  $password = Find-SEElement -Name 'passwd'
  $submitBtn = Find-SEElement -ClassName 'btn'
  Send-SEKeys -Element $username -Keys 'USERNAME'
  Send-SEKeys -Element $password -Keys 'PASSWORD'
  Invoke-SEClick -Element $submitBtn

  .Example
  # Login to Reddit.com using pipeline
  Find-SEElement -Name 'user' | Send-SEKeys -Keys 'USERNAME'
  Find-SEElement -Name 'passwd' | Send-SEKeys -Keys 'PASSWORD'
  Find-SEElement -ClassName 'btn' | Invoke-SEClick

  .Notes
#>
function Send-SEKeys {
    param(
    [Parameter(Position=0,Mandatory=$true,ValueFromPipeline=$true)]
    [OpenQA.Selenium.IWebElement]$Element,
    [Parameter(Mandatory=$true)]
    [string]$Keys,
    [switch]$SubmitForm
    )

    $Element.SendKeys($Keys)
    if($SubmitForm -eq $true){
        try{
            $Element.Submit()
        }
        catch{
            throw "Cannot submit, this is not a form."
        }
    }
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

<#
 .Synopsis
  Gets a Selenium Web Element Attribute.

 .Description
  This command will allow you to find the value of an attribute on your
  Selenium Web Element.

 .Parameter Element
  Selenium Element Object

 .Parameter Attribute
  Specify the Attribute you want from the Element.

  .Notes
#>
function Get-SEElementAttribute {
    param(
        [Parameter(Position=0,ValueFromPipeline=$true,Mandatory = $true)]
        [OpenQA.Selenium.IWebElement]$Element,
        [Parameter(Mandatory=$true)]
        [string]$Attribute
    )

    Process {
        $Element.GetAttribute($Attribute)
    }
}