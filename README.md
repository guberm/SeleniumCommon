# SeleniumCommon

A .NET library providing common utilities and helper methods for Selenium WebDriver automation testing.

## 📋 Overview

SeleniumCommon is a utility library designed to simplify and standardize Selenium WebDriver automation tasks. It provides reusable components for web testing, including driver management, screenshot capture, PDF reading, network connections, and comprehensive reporting capabilities.

## 🚀 Features

- **WebDriver Management**: Automated Chrome driver setup and configuration
- **Screenshot Utilities**: Capture screenshots for test failures and debugging
- **PDF Reading**: Extract text content from PDF files using iTextSharp
- **Database Integration**: Helper methods for database operations in test environments
- **Extent Reports**: Advanced HTML reporting with test results and screenshots
- **Network Utilities**: Handle network drive connections with authentication
- **Error Handling**: Comprehensive error detection and reporting for web applications
- **Alert Management**: Automated handling of browser alerts and dialogs

## 🛠️ Technologies

- **.NET Framework 4.7.2**
- **Selenium WebDriver 4.15.0**
- **NUnit 3.14.0** for testing framework
- **ExtentReports 4.0.3** for test reporting
- **iTextSharp 5.5.10** for PDF processing
- **log4net 2.0.10** for logging

## 📦 Installation

1. Clone the repository:
```bash
git clone https://github.com/guberm/SeleniumCommon.git
```

2. Restore NuGet packages:
```bash
nuget restore SeleniumCommon.sln
```

3. Build the solution:
```bash
msbuild SeleniumCommon.sln
```

## ⚙️ Configuration

Before using the library, update the `App.config` file with your specific settings:

```xml
<appSettings>
    <add key="correctTramsEndpoint" value="your-correct-endpoint:24002/TRAMSService.svc/v2" />
    <add key="incorrectTramsEndpoint" value="your-incorrect-endpoint:24009/TRAMSService.svc/v2" />
    <add key="url" value="your-test-url" />
    <add key="username" value="your-username" />
    <add key="is_trams_up" value="true" />
</appSettings>

<connectionStrings>
    <add name="AutomationDB" connectionString="Data Source=your-server;Initial Catalog=AutomationDB;Integrated Security=true;" />
    <add name="TravelEdgeDB" connectionString="Data Source=your-server;Initial Catalog=TravelEdge;Integrated Security=true;" />
</connectionStrings>
```

## 🔧 Usage Examples

### Basic WebDriver Setup
```csharp
using SeleniumCommon;

// Start Chrome driver with automatic configuration
IWebDriver driver = Common.StartChromeDriver("https://example.com");

// Perform your automation tasks
// ...

// Clean up
Common.EndDriver(driver);
```

### Screenshot Capture
```csharp
// Capture screenshot on test failure
string screenshotPath = Common.GetScreenShot(driver, "test_failure_screenshot");
```

### PDF Text Extraction
```csharp
var pdfReader = new PdfReader();
string textContent = pdfReader.GetTextFromPdf("path/to/document.pdf");
```

### Database Operations
```csharp
// Load configuration from database
Common.LoadConfiguration("configuration_table");

// Get data from database
string result = Common.GetDataFromDb("SELECT value FROM settings WHERE key = 'test_setting'");
```

### Error Detection
```csharp
// Check for application errors on the page
bool hasErrors = Common.ErrorsChecker(driver);
if (hasErrors) {
    // Handle error condition
}
```

## 📊 Reporting

The library includes integrated ExtentReports functionality:

```csharp
public class MyTestClass : ExtentReport
{
    [Test]
    public void MyTest()
    {
        test = extent.CreateTest("My Test Case");
        
        // Your test logic here
        
        // Results are automatically captured
        getResult(driver);
    }
}
```

## 🔒 Security Notes

⚠️ **Important Security Considerations:**

- All sensitive data (database connections, encryption keys, server credentials) should be stored securely
- Never commit actual credentials to version control
- Use environment variables or secure configuration management for production deployments
- Regularly update dependencies to patch security vulnerabilities

## 🐛 Known Issues

- Chrome driver version compatibility requires periodic updates
- Hard-coded file paths may need adjustment for different environments
- Some legacy code patterns may need refactoring for modern .NET versions

## 🤝 Contributing

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/amazing-feature`)
3. Commit your changes (`git commit -m 'Add amazing feature'`)
4. Push to the branch (`git push origin feature/amazing-feature`)
5. Open a Pull Request

## 📝 Recent Changes

See [CODE_REVIEW_FIXES.md](CODE_REVIEW_FIXES.md) for detailed information about recent code improvements and security fixes.

## 📄 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## 🆘 Support

For questions, issues, or contributions, please:
- Open an issue on GitHub
- Contact the repository maintainer
- Check the existing documentation and code comments

## 🔄 Version History

- **v1.0.0** - Initial release with core Selenium utilities
- **v1.1.0** - Added security improvements and namespace standardization
- **v1.2.0** - Updated dependencies and improved error handling

---

**Note**: This library is designed for test automation purposes. Ensure proper configuration and security measures when using in production environments.
