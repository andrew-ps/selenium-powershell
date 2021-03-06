# Selenium PowerShell Module

- Wraps the C# WebDriver for Selenium
- Easily execute web-based tests

## Example Script
This will open up reddit.com in Chrome, login, and open the rising posts tab.
```PowerShell
Import-Module .\Documents\Source\selenium-powershell\Selenium.psd1 -Force
# This is how you would launch Chrome with special switches
$driverOptions = New-Object OpenQA.Selenium.Chrome.ChromeOptions

# List of switches https://peter.sh/experiments/chromium-command-line-switches/
$driverOptions.AddArguments("disable-infobars")

# This would disable the prompt to save credentials on logins
$driverOptions.AddUserProfilePreference("credentials_enable_service", $false)

# Create driver and navigate to Reddit
$driver = Start-SEChrome -Url 'http://reddit.com' -Options $driverOptions -LeaveBrowserRunning

# Login to Reddit
$usernameField = Find-SEElement -Driver $driver -Name 'user'
Send-SEKeys -Element $usernameField -Keys 'USERNAME'
Find-SEElement -Driver $driver -Name 'passwd' | Send-SEKeys -Keys 'PASSWORD'
Find-SEElement -Driver $driver -ClassName 'btn' | Invoke-SEClick

# Wait for webpage to load before moving on by checking for an element that
# only exists on the next page
# This is generally not needed unless Selenium moves onto next command before
# completing the previous action. Usually when dealing with a javascript form
Invoke-SEWait -Driver $driver -ClassName 'userkarma' -Seconds 10

# Check for new mail by looking at class attribute
$mail = Find-SEElement $driver -Id 'mail'
Get-SEElementAttribute -Element $mail -Attribute 'class'

# Display rising reddit posts
$risingPosts = Find-SEElement -Driver $driver -LinkText 'rising'
Invoke-SEClick -Element $risingPosts

# Take Screenshot of rising posts and send it as an email attachment
$screenshot = Get-SEScreenshot -Driver $driver -Format Bmp -AsMemoryStream
$attachment = [System.Net.Mail.Attachment]::new($screenshot,'risingposts.bmp')
$message = [System.Net.Mail.MailMessage]::new('example@myorg.com','example@myorg.com')
$message.Attachments.Add($attachment)
$smtp = [System.Net.Mail.SmtpClient]::new()
$smtp.Host = 'smtp.office365.com'
$username = 'example@myorg.com'
$password = 'Password'
$smtp.Credentials = [System.Net.NetworkCredential]::new($username,$password)
$smtp.Port = 587
$smtp.EnableSsl = $true
$smtp.Send($message)

# Do a reddit search
Find-SEElement -Driver $Driver -Name 'q' | Send-SEKeys -Keys 'Test Search' -SubmitForm
```

## Driver Versions

- Chrome Driver : 2.31.488763
- Edge Driver : 5.16299
  * Requires Windows 10 v 1709. Download other versions [here](https://developer.microsoft.com/en-us/microsoft-edge/tools/webdriver/)
- Gecko Driver : 0.19.1

## Notes
There is a function for the Internet Explorer driver but it has not been tested due to all the prerequisites required before using. The prerequisites can be found [here](https://github.com/SeleniumHQ/selenium/wiki/InternetExplorerDriver#required-configuration) and the driver can be found on the same page as well.

# Using Selenium Without this Module
Incase you are looking to do things not available in this module or just want to see how to use Selenium without this module here is the same exact sample as above without this module's functions.
```PowerShell
# Add the Selenium Dlls
Add-Type -Path "C:\mydir\WebDriver.dll"
Add-Type -Path "C:\mydir\WebDriver.Support.dll"
Add-Type -Path "C:\mydir\Selenium.WebDriverBackedSelenium.dll"
# This is how you would launch Chrome with special switches
$driverOptions = New-Object OpenQA.Selenium.Chrome.ChromeOptions

# List of switches https://peter.sh/experiments/chromium-command-line-switches/
$driverOptions.AddArguments("disable-infobars")

# This should keep the browser running after the driver closes but does not seem to work from console
$driverOptions.LeaveBrowserRunning = $true

# This would disable the prompt to save credentials on logins
$driverOptions.AddUserProfilePreference("credentials_enable_service", $false)

# Create driver and navigate to Reddit
$driver = New-Object OpenQA.Selenium.Chrome.ChromeDriver($driverOptions)
$driver.Navigate().GoToUrl("http://reddit.com")

# Login to Reddit
$usernameField = $driver.FindElementByName('user')
$passwordField = $driver.FindElementByName('passwd')
$submitBtn = $driver.FindElementByClassName('btn')
$usernameField.SendKeys('USERNAME')
$passwordField.SendKeys('PASSWORD')
$submitBtn.Click()

# Wait for webpage to load before moving on by checking for an element that
# only exists on the next page
# This is generally not needed unless Selenium moves onto next command before
# completing the previous action. Usually when dealing with a javascript form
$wait = [OpenQA.Selenium.Support.UI.WebDriverWait]::new($driver,10)
$wait.Until([OpenQA.Selenium.Support.UI.ExpectedConditions]::ElementExists([OpenQA.Selenium.By]::ClassName('userkarma')))

# Check for new mail by looking at class attribute
$mail = $driver.FindElementById('mail')
$mail.GetAttribute('class')

# Display rising reddit posts
$risingPosts = $driver.FindElementByLinkText('rising')
$risingPosts.Click()

# Take Screenshot of rising posts and send it as an email attachment
$screenshot = $driver.GetScreenshot() # The default screenshot format is PNG
$screenshotStream = [System.IO.MemoryStream]::new($screenshot.AsByteArray) # Save screenshot in memorystream variable
$screenshotBitmap = [System.Drawing.Bitmap]::new($screenshotStream) # Create new bitmap of screenshot
$formattedStream = [System.IO.MemoryStream]::new() # Create new memorystream to save formatted screenshot to
$screenshotBitmap.Save($formattedStream,[System.Drawing.Imaging.ImageFormat]::Bmp) # Save image to memorystream in Bmp format
$screenshotStream = $formattedStream
$screenshotStream.Position = 0 # Set stream position to zero otherwise attachment will be invalid
$attachment = [System.Net.Mail.Attachment]::new($screenshotStream,'risingposts.bmp')
$message = [System.Net.Mail.MailMessage]::new('example@myorg.com','example@myorg.com')
$message.Attachments.Add($attachment)
$smtp = [System.Net.Mail.SmtpClient]::new()
$smtp.Host = 'smtp.office365.com'
$username = 'example@myorg.com'
$password = 'Password'
$smtp.Credentials = [System.Net.NetworkCredential]::new($username,$password)
$smtp.Port = 587
$smtp.EnableSsl = $true
$smtp.Send($message)

# Do a reddit search
$search = $driver.FindElementByName('q')
$search.SendKeys('Test Search')
$search.Submit()
```
