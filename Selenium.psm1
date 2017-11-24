[System.Reflection.Assembly]::LoadFrom("$PSScriptRoot\WebDriver.dll")
[System.Reflection.Assembly]::LoadFrom("$PSScriptRoot\WebDriver.Support.dll")
[System.Reflection.Assembly]::LoadFrom("$PSScriptRoot\Selenium.WebDriverBackedSelenium.dll")

function Start-SEChrome {
    param(
    [Parameter(Mandatory=$false)]
    [OpenQA.Selenium.Chrome.ChromeOptions]$Options
    )
    if($options -ne $null){
        New-Object OpenQA.Selenium.Chrome.ChromeDriver($Options)
    }
    else{
        New-Object -TypeName "OpenQA.Selenium.Chrome.ChromeDriver"
    }
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
  Selenium Element Object

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
        [Parameter(ParameterSetName = "DriverByName")]
        [Parameter(ParameterSetName = "DriverById")]
        [Parameter(ParameterSetName = "DriverByClassName")]
        [Parameter(ParameterSetName = "DriverByLinkText")]
        [Parameter(ParameterSetName = "DriverByTagName")]
        [Parameter(ParameterSetName = "DriverByPartialLinkText")]
        [Parameter(ParameterSetName = "DriverByCssSelector")]
        [Parameter(ParameterSetName = "DriverByXPath")]
        [OpenQA.Selenium.Remote.RemoteWebDriver]$Driver,
        [Parameter(ParameterSetName = "ElementByName")]
        [Parameter(ParameterSetName = "ElementById")]
        [Parameter(ParameterSetName = "ElementByClassName")]
        [Parameter(ParameterSetName = "ElementByLinkText")]
        [Parameter(ParameterSetName = "ElementByTagName")]
        [Parameter(ParameterSetName = "ElementByPartialLinkText")]
        [Parameter(ParameterSetName = "ElementByCssSelector")]
        [Parameter(ParameterSetName = "ElementByXPath")]
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
                $Target.FindElements([OpenQA.Selenium.By]::LinkText($LinkText))
            }
            catch{
                throw "No Element Found"
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
    [Parameter(Mandatory=$true,ValueFromPipeline=$true)]
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
    [Parameter(Mandatory=$true,ValueFromPipeline=$true)]
    [OpenQA.Selenium.IWebElement]$Element,
    [string]$Keys
    )

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

 .Example
  # Get the Reddit.com Search Bar Element and Input a String
  # HTML EXAMPLE
  $searchBar = Get-SEElement -Driver $driver -FindBy Name -Value 'q'
  $searchBar.SendKeys('PowerCLI Module')

 .Example
  # Get a Table Element and then Find a SubElement in the Table
  $table = Get-SEElement -Driver $driver -FindBy ## -Value ##
  $rowTwo = Get-SEElement -Element $table -FindBy ## -Value ## 

  .Notes
#>
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