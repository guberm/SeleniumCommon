using Automation.Test;
using AventStack.ExtentReports.Utils;
using log4net;
using Microsoft.Expression.Encoder.ScreenCapture;
using NUnit.Framework;
using NUnit.Framework.Interfaces;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Remote;
using OpenQA.Selenium.Support.UI;
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Configuration;
using System.Data.SqlClient;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.IO.Compression;
using System.Linq;
using System.Net;
using System.Reflection;
using System.Security.Cryptography;
using System.Text;
using System.Threading;
using System.Xml;
using Microsoft.Win32;

namespace Automation
{
    public class Common : ExtentReport
    {
        private static readonly ILog Logger =
            LogManager.GetLogger(MethodBase.GetCurrentMethod().DeclaringType);

        private static string _env;
        private static string connString;

        private static ScreenCaptureJob _videorec;
        private static TestStatus _testStatus;
        private const string InitVector = "xxx";
        private const int Keysize = 256;
        private static string passPhrase = "yyy";

        public static ConcurrentDictionary<string, string> ModuleData;
        //private static string videoFile;

        public static bool IsTextExist(IWebDriver driver, string text) {
            AcceptAlert(driver);
            bool isTextExist = GetPageContent(driver).Contains(text);
            return isTextExist;
        }

        public static string GetPageContent(IWebDriver driver) {
            AcceptAlert(driver);

            string pageSource = driver.PageSource;

            return pageSource;
        }

        public static void AcceptAlert(IWebDriver driver) {
            try
            {
                IAlert alert = ExpectedConditions.AlertIsPresent().Invoke(driver);
                if (alert == null) return;
                alert = driver.SwitchTo().Alert();
                alert.Accept();
            }
            catch { }
        }

        public static bool IsTextDisplayed(IWebDriver driver, string text) {
            AcceptAlert(driver);
            bool isTextDisplayed = false;
            driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(2);
            try {
                IWebElement page = driver.FindElement(By.XPath(".//*"));
                isTextDisplayed = page.Text.Contains(text);
            }
            catch {
                isTextDisplayed = false;
            }

            driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromMinutes(5);
            return isTextDisplayed;
        }

        public static string GetScreenShot(IWebDriver driver, string screenShotName) {
            string path = "C:\\AutomationErrorScreenshots\\" + DateTime.Now.ToString("yyyy-MM-dd") + "\\";
            bool exists = Directory.Exists(path);
            if (!exists) {
                Directory.CreateDirectory(path);
            }

            try {
                ITakesScreenshot ts = (ITakesScreenshot) driver;
                Screenshot screenshots = ts.GetScreenshot();
                //string pth = System.Reflection.Assembly.GetCallingAssembly().CodeBase;
                //string finalpth = pth.Substring(0, pth.LastIndexOf("bin")) + "ErrorScreenshots\\" + screenShotName + ".png";
                string finalpth = path + screenShotName + ".png";
                string localpath = new Uri(finalpth).LocalPath;
                screenshots.SaveAsFile(localpath, ScreenshotImageFormat.Png);
                return localpath;
            }
            catch {
                return null;
            }
        }

        public static void GetWindowScreenShot(string screenShotName) {
            string path = "C:\\AutomationErrorScreenshots\\" + DateTime.Now.ToString("yyyy-MM-dd") + "\\";
            bool exists = Directory.Exists(path);
            if (!exists) {
                Directory.CreateDirectory(path);
            }

            // Start the process...
            Bitmap memoryImage = new Bitmap(1915, 1000);
            Size s = new Size(memoryImage.Width, memoryImage.Height);

            // Create graphics
            Graphics memoryGraphics = Graphics.FromImage(memoryImage);

            // Copy data from screen
            memoryGraphics.CopyFromScreen(0, 0, 0, 0, s);

            //That's it! Save the image in the directory and this will work like charm.
            string finalpth = null;
            try {
                //string pth = System.IO.Path
                //    .GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase)
                //    ?.Substring(6);
                //finalpth = pth.Substring(0, pth.LastIndexOf("bin")) + "ErrorScreenshots\\" + screenShotName + ".png";

                finalpth = path + screenShotName + ".png";
            }
            catch { }

            // Save it!
            if (finalpth != null) memoryImage.Save(finalpth, ImageFormat.Png);
        }

        public static void LoadConfiguration(string dbTableName) {
            // skip if already loaded
            //if (ModuleData == null)
            //{
            // clear config
            ModuleData = new ConcurrentDictionary<string, string>();

            // load connection string
            //connString = ConfigurationManager.ConnectionStrings["AutomationDB"].ConnectionString;

            // connect to db
            using (SqlConnection conn =
                new SqlConnection(ConfigurationManager.ConnectionStrings["AutomationDB"].ConnectionString)) {
                // open connection
                conn.Open();

                // prepare read query
                SqlCommand command = new SqlCommand($"SELECT [name], [value] FROM [{dbTableName}]", conn);

                // execute
                SqlDataReader configReader = command.ExecuteReader();
                try {
                    // iterate results
                    while (configReader.Read()) {
                        // add to config dict
                        ModuleData.TryAdd(
                            configReader["name"].ToString(),
                            configReader["value"].ToString());
                    }
                }
                finally {
                    // close reader
                    configReader.Close();
                }

                conn.Close();
            }

            //}
        }

        public static bool IsElementClicked(IWebDriver driver, IWebElement element) {
            driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(2);
            try {
                try {
                    element.Click();
                    driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromMinutes(1);
                    return true;
                }
                catch {
                    driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromMinutes(1);
                    return false;
                }
            }
            catch (NoSuchElementException) {
                driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromMinutes(1);
                return false;
            }
        }

        public static bool IsElementPresent(IWebDriver driver, IWebElement element) {
            driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(2);
            try {
                string tag = element.TagName;
                return true;
            }
            catch {
                driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromMinutes(1);
                return false;
            }
        }

        public static bool IsElementPresent(IWebDriver driver, By by) {
            driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(2);
            try {
                ReadOnlyCollection<IWebElement> uuu = driver.FindElements(by);
                return uuu.Count > 0;
            }
            catch (NoSuchElementException) {
                driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromMinutes(2);
                return false;
            }
        }

        public static bool IsElementDisplayed(IWebDriver driver, By by) {
            driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(2);
            try {
                bool uuu = driver.FindElement(by).Displayed;
                return uuu;
            }
            catch (NoSuchElementException) {
                driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromMinutes(2);
                return false;
            }
        }

        public static bool IsElementDisplayed(IWebDriver driver, IWebElement element) {
            driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(2);
            try {
                bool uuu = element.Displayed;
                return uuu;
            }
            catch (NoSuchElementException) {
                driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromMinutes(2);
                return false;
            }
        }

        public static bool ErrorsChecker(IWebDriver driver) {
            List<string> errorsMsgs = new List<string> {
                "Send Support Request",
                "Developer Message",
                "An error has occurred",
                "An error occurred",
                //"Our backend system is unavailable",
                "Room unavailable"
            };

            string pageContent = GetPageContent(driver);

            int errorId = 0;
            int errorTypeId = 0;
            string errorName = null;
            bool toContinue = false;

            if (pageContent.ToLower().Contains("trams is currently inaccessible")) {
                return true;
            }

            foreach (string errorsMsg in errorsMsgs) {
                if (pageContent.ToLower().Contains(errorsMsg.ToLower())) {
                    if (errorsMsg == "send support request" || errorsMsg == "developer message") {
                        toContinue = false;

                        string error = driver.FindElement(By.XPath("//*[contains(@class,'error-guid')]")).Text;

                        errorId = Convert.ToInt32(error.Substring(0, error.IndexOf(':')));

                        errorName = GetDataFromDb(
                            $"SELECT ExceptionFullName FROM [dbo].[ErrorMessages] WHERE ErrorMessageId = {errorId}");
                        errorTypeId = Convert.ToInt32(GetDataFromDb(
                            $"SELECT ErrorTypeId FROM [dbo].[ErrorMessages] WHERE ErrorMessageId = {errorId}"));
                    }
                    else if (errorName == null) {
                        Logger.Error(
                            $"{TestContext.CurrentContext.Test.FullName} - Test failed: '{errorsMsg}' error received");
                        Assert.Fail($"Process failed: '{errorsMsg}' error received");
                    }
                    else if (pageContent.ToLower().Contains("please sign in to continue")) {
                        toContinue = true;

                        driver.FindElement(By.XPath("//*[@id='password']")).SendKeys("Pass1234");
                        Thread.Sleep(500);
                        driver.FindElement(By.XPath("//button[.='Sign in']")).Click();
                        Thread.Sleep(500);
                    }
                    else if (pageContent.ToLower().Contains("air itinerary price change")) {
                        toContinue = true;

                        driver.FindElement(By.XPath("//*[@value='new-price']")).Click();
                        Thread.Sleep(500);
                        driver.FindElement(By.XPath("//button[.='Continue']")).Click();
                        Thread.Sleep(500);
                    }
                    else {
                        if (errorTypeId == 1) {
                            toContinue = false;

                            Logger.Error(
                                $"{TestContext.CurrentContext.Test.FullName} - Test failed: '{errorsMsg}' error received, error Id = {errorId}, errorName = {errorName}");
                            Assert.Fail(
                                $"Process failed: '{errorsMsg}' error received, error Id = {errorId}, errorName = {errorName}");
                        }
                        else {
                            toContinue = false;

                            Logger.Error(
                                $"{TestContext.CurrentContext.Test.FullName} - Test failed: '{errorsMsg}' error received, error Id = {errorId}, errorName = {errorName}");
                            Assert.Fail(
                                $"Process failed: '{errorsMsg}' error received, error Id = {errorId}, errorName = {errorName}");
                        }
                    }
                }
            }

            if (pageContent.ToLower().Contains("our backend system is unavailable")) {
                return true;
            }

            if (pageContent.ToLower().Contains("selected flights &amp; classes are no longer available")) {
                toContinue = false;

                Actions actions = new Actions(driver);
                actions.SendKeys(Keys.End);
                actions.Perform();

                Logger.Error(
                    $"{TestContext.CurrentContext.Test.FullName} - Test failed: Selected flights & classes are no longer available");
                Assert.Fail("Process failed: Selected flights & classes are no longer available");
            }

            if (pageContent.ToLower().Contains("<div class=\"dashboard-title-text\">error</div>")) {
                toContinue = false;

                Logger.Error($"{TestContext.CurrentContext.Test.FullName} - Test failed: An error occured");
                Assert.Fail("Process failed: An error occurred");
            }

            return toContinue;
        }

        public static string GetDataFromDb(string query, string db = null) {
            string value = null;
            connString = db == "AutomationDB"
                ? ConfigurationManager.ConnectionStrings["AutomationDB"].ConnectionString
                : $"data source=10.10.2.116;initial catalog=TravelEdge_{_env};persist security info=true;user id=traveledge;password=qqq;multipleactiveresultsets=true";

            using (SqlConnection conn =
                new SqlConnection(connString)) {
                conn.Open();
                SqlCommand request = new SqlCommand(query, conn);
                using (SqlDataReader reader = request.ExecuteReader()) {
                    if (reader.Read()) {
                        value = reader.GetValue(0).ToString();
                    }
                }

                conn.Close();
            }

            return value;
        }

        public static Dictionary<string, string> GetDataFromDbDictionary(string query, string db = null) {
            Dictionary<string, string> dbData = new Dictionary<string, string>();

            connString = db == "AutomationDB"
                ? ConfigurationManager.ConnectionStrings["AutomationDB"].ConnectionString
                : $"data source=10.10.2.116;initial catalog=TravelEdge_{_env};persist security info=true;user id=traveledge;password=qqq;multipleactiveresultsets=true";

            using (SqlConnection conn = new SqlConnection(connString)) {
                conn.Open();
                SqlCommand request = new SqlCommand(query, conn);
                using (SqlDataReader reader = request.ExecuteReader()) {
                    while (reader.Read()) {
                        dbData.Add(reader.GetValue(0).ToString(), reader.GetValue(1).ToString());
                    }
                }

                conn.Close();
            }

            return dbData;
        }

        public static int WriteDataToDb(string insertCommand, string db) {
            int affectedRows;

            connString = db == "AutomationDB"
                ? ConfigurationManager.ConnectionStrings["AutomationDB"].ConnectionString
                : $"data source=10.10.2.116;initial catalog=TravelEdge_{_env};persist security info=true;user id=traveledge;password=qqq;multipleactiveresultsets=true";

            using (SqlConnection conn =
                new SqlConnection(connString)) {
                conn.Open();
                using (SqlCommand cmd = new SqlCommand(insertCommand, conn)) {
                    affectedRows = cmd.ExecuteNonQuery();
                }

                conn.Close();
            }

            return affectedRows;
        }

        public static void DeleteDataFromDb(string deleteCommand, string db) {
            connString = db == "AutomationDB"
                ? ConfigurationManager.ConnectionStrings["AutomationDB"].ConnectionString
                : $"data source=10.10.2.116;initial catalog=TravelEdge_{_env};persist security info=true;user id=traveledge;password=qqq;multipleactiveresultsets=true";

            using (SqlConnection conn =
                new SqlConnection(connString)) {
                conn.Open();
                using (SqlCommand command = new SqlCommand(deleteCommand, conn)) {
                    command.ExecuteNonQuery();
                }

                conn.Close();
            }
        }

        public static IWebDriver StartChromeDriver(string url) {
            Retry:

            _env = GetEnvironment(url);

            connString =
                $"data source=10.10.2.116;initial catalog=TravelEdge_{_env};persist security info=true;user id=traveledge;password=qqq;multipleactiveresultsets=true"; //ConfigurationManager.ConnectionStrings[db].ConnectionString;

            log4net.Config.XmlConfigurator.Configure();

            string min = DateTime.Now.Minute.ToString();
            if (min.Length == 1) {
                min = $"0{min}";
            }

            string hour = DateTime.Now.Hour.ToString();
            if (hour.Length == 1) {
                hour = $"0{hour}";
            }

            string testName = TestContext.CurrentContext.Test.ClassName.Replace("Automation.Tests.", "");
            GlobalContext.Properties["LogName"] =
                $"Automation_{_env}_{testName}.{TestContext.CurrentContext.Test.Name}_{hour}-{min}.log";
            GlobalContext.Properties["ApplicationName"] =
                Assembly.GetExecutingAssembly().GetName().Name;
            string pth = Assembly.GetExecutingAssembly().Location.Replace("Automation.dll", "");
            FileInfo logfile = new FileInfo($"{pth}\\Automation.dll.config");
            log4net.Config.XmlConfigurator.ConfigureAndWatch(logfile);

            if (!Convert.ToBoolean(ConfigurationManager.AppSettings["is_trams_up"])) {
                UpdateTramServer(false);
            }

            TimeSpan timeout = TimeSpan.FromMinutes(5);
            IWebDriver driver = null;
            try {
                ChromeDriverService chromeService = ChromeDriverService.CreateDefaultService($"{pth}");

                ChromeOptions chromeOptions = new ChromeOptions();
                chromeOptions.AddArgument("start-maximized");
                chromeOptions.AddArgument("no-sandbox");
                //chromeOptions.AddExtension($"{pth}Resources\\ChroPath_v5.0.5.crx");

                chromeOptions.AddAdditionalCapability(CapabilityType.EnableProfiling, true, true);
                chromeOptions.SetLoggingPreference(LogType.Server, LogLevel.All);

                chromeOptions.AddUserProfilePreference("download.default_directory",
                    @"C:\AutomationItinerary\" + DateTime.Now.ToString("yyyy-MM-dd"));

                driver = new ChromeDriver(chromeService, chromeOptions, timeout);
            }
            catch (Exception e) {
                if (e.Message.Contains("SessionNotCreated") || e.Message.Contains("chromedriver.exe does not exist")) {
                    try {
                        string path = Assembly.GetExecutingAssembly().Location;
                        string ppp = path.Split('\\').Last();
                        path = path.Replace(ppp, "");

                        CommandLine($"{path}Resources\\KillChromeDriver.exe");
                    }
                    catch { }

                    GetLatestChromeDriver();

                    goto Retry;
                }
            }

            driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromMinutes(5);
            driver.Manage().Timeouts().AsynchronousJavaScript = TimeSpan.FromMinutes(5);
            driver.Manage().Timeouts().PageLoad = TimeSpan.FromMinutes(5);

            driver.Url = url;
            driver.Manage().Cookies.DeleteAllCookies();

            Logger.Info(TestContext.CurrentContext.Test.FullName + " - Test started");

            //videoFile = StartRecordVideo();

            return driver;
        }

        public static void GetLatestChromeDriver()
        {
            string chromeVer = (string)Registry.GetValue(@"HKEY_CURRENT_USER\Software\Google\Chrome\BLBeacon", "version", null);

            chromeVer = chromeVer.Split('.')[0];

            string path = Assembly.GetExecutingAssembly().Location;
            string ppp = path.Split('\\').Last();
            path = path.Replace(ppp, "");

            if (File.Exists($"{path}chromedriver.exe"))
            {
                File.Delete($"{path}chromedriver.exe");
            }

            string ver = GetHtmlDocument("https://chromedriver.storage.googleapis.com/");

            List<string> versions = new List<string>();

            XmlDocument xml = new XmlDocument();
            xml.LoadXml(ver);
            foreach (XmlNode node in xml.DocumentElement.ChildNodes)
            {
                switch (node.Name)
                {
                    case "Contents":
                        {
                            foreach (XmlElement xmlElement in node.Cast<XmlElement>().Where(xmlElement => xmlElement.Name == "Key"))
                            {
                                try
                                {
                                    if (xmlElement.InnerText.Contains($"LATEST_RELEASE_{chromeVer}"))
                                    {
                                        versions.Add(xmlElement.InnerText);
                                    }
                                }
                                catch
                                {
                                    // ignored
                                }
                            }

                            break;
                        }
                }
            }

            ver = GetHtmlDocument($"https://chromedriver.storage.googleapis.com/{versions.Last()}");

            using (WebClient webClient = new WebClient())
            {
                webClient.DownloadFile($"https://chromedriver.storage.googleapis.com/{ver}/chromedriver_win32.zip",
                    $"{path}chromedriver.zip");
            }

            int count = 0;
            while (!File.Exists($"{path}chromedriver.zip") && count < 10)
            {
                Thread.Sleep(1000);
            }

            if (File.Exists($"{path}chromedriver.zip"))
            {
                ZipFile.ExtractToDirectory($"{path}chromedriver.zip", $"{path}");
            }

            if (File.Exists($"{path}chromedriver.zip"))
            {
                File.Delete($"{path}chromedriver.zip");
            }
        }

        private static string GetHtmlDocument(string url) {
            HttpWebRequest request = (HttpWebRequest) WebRequest.Create(url);
            request.MaximumAutomaticRedirections = 300;
            request.MaximumResponseHeadersLength = 300;
            request.AllowWriteStreamBuffering = true;
            request.Credentials = CredentialCache.DefaultCredentials;
            HttpWebResponse response = (HttpWebResponse) request.GetResponse();
            Stream receiveStream = response.GetResponseStream();
            StreamReader readStream = new StreamReader(receiveStream, Encoding.UTF8);
            string resp = readStream.ReadToEnd();
            response.Close();
            readStream.Close();

            return resp;
        }

        public static void EndDriver(IWebDriver driver) {
            //videorec.Stop();
            _testStatus = TestContext.CurrentContext.Result.Outcome.Status;
            if (_testStatus != TestStatus.Passed) {
                if (TestContext.CurrentContext.Result.Message.Contains("timed out after 60 seconds") ||
                    TestContext.CurrentContext.Result.Message.Contains("{Alert text :")) {
                    GetWindowScreenShot(TestContext.CurrentContext.Test.FullName + "_" +
                                        DateTime.Now.ToString("yyyy-MM-ddTHH:mm:ss").Replace(':', '_'));
                }
                else {
                    GetScreenShot(driver,
                        TestContext.CurrentContext.Test.FullName + "_" +
                        DateTime.Now.ToString("yyyy-MM-ddTHH:mm:ss").Replace(':', '_'));
                }

                Logger.Error($"URL: {driver.Url}");

                driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(1);
                try {
                    Logger.Error($"GUID: {driver.FindElement(By.XPath("//*[contains(@class,'error-guid')]")).Text}");
                }
                catch { }

                Logger.Error($"{TestContext.CurrentContext.Test.FullName} - Test Failed");
                Logger.Error(TestContext.CurrentContext.Result.Message);
                Logger.Error(TestContext.CurrentContext.Result.StackTrace);
            }
            else {
                Logger.Info(TestContext.CurrentContext.Test.FullName + " - Test Finished");
                //File.Delete(videoFile);
            }

            driver.Quit();
        }

        public static void IsWaitMsgPresented(IWebDriver driver) {
            string cUrl1 = driver.Url;

            Thread.Sleep(2000);

            int count = 0;
            while (IsTextDisplayed(driver, "Please wait while we process your request") && count < 300) {
                Thread.Sleep(1000);
                ErrorsChecker(driver);
                count++;

                if (Common.IsElementPresent(driver,
                    By.XPath("//*[contains(text(),'The cost of the Insurance has been')]"))) {
                    IWebElement continueBtnNotif =
                        driver.FindElement(By.XPath("//button[contains(text(),'Continue')]"));
                    continueBtnNotif.Click();
                }
            }

            string cUrl2 = driver.Url;
            if (cUrl1 == cUrl2) {
                if (IsTextDisplayed(driver, "Please wait while we process your request") && count > 0) {
                    Assert.Fail("Waiting message continues for than 5 minutes");
                }
            }

            Thread.Sleep(500);
        }

        public static void IsWaitFlightMsgPresented(IWebDriver driver) {
            int count = 0;
            while (IsTextDisplayed(driver, "Please wait while we fetch your flight results") && count < 300) {
                Thread.Sleep(1000);
                ErrorsChecker(driver);
                count++;
            }

            Thread.Sleep(500);
        }

        public static void PriceWaitingMsg(IWebDriver driver) {
            int count = 0;
            while (IsTextExist(driver, "Please wait while we fetch your pricing results.") && count < 300) {
                Thread.Sleep(1000);
                ErrorsChecker(driver);
                count++;

                ErrorsChecker(driver);
            }

            Thread.Sleep(1000);
        }

        public static void TravelersWaitingMsg(IWebDriver driver) {
            int count = 0;

            IAlert alert;
            while (IsTextDisplayed(driver, "Please wait while travelers are updated") && count < 300) {
                ErrorsChecker(driver);
                Thread.Sleep(500);
                count++;
                ErrorsChecker(driver);
                AcceptAlert(driver);

                ErrorsChecker(driver);
            }

            Thread.Sleep(1000);

            ErrorsChecker(driver);
            AcceptAlert(driver);

            ErrorsChecker(driver);
        }

        public static void ClickOnElement(IWebDriver driver, IWebElement element) {
            Actions actions = new Actions(driver);
            int countHome = 1;
            int count = 0;

            ErrorsChecker(driver);

            while (!IsElementClicked(element, driver) && count < 100) {
                ErrorsChecker(driver);
                IWebElement el = driver.FindElements(By.XPath("//*[@class='row']")).First();

                if (count == 0) {
                    actions.MoveToElement(el).Click().Perform();
                    Thread.Sleep(500);
                }

                actions.SendKeys(Keys.ArrowDown).Click();
                actions.Perform();
                Thread.Sleep(500);

                ErrorsChecker(driver);

                count++;

                if (countHome == 1) {
                    Actions actions1 = new Actions(driver);
                    actions1.SendKeys(Keys.Home).Perform();
                    Thread.Sleep(500);

                    countHome++;
                    ErrorsChecker(driver);
                }

                ErrorsChecker(driver);
            }

            if (count > 99) {
                Assert.Fail($"Element \"{element}\" is not clickable or not available");
            }
        }

        private static bool IsElementClicked(IWebElement element, IWebDriver driver) {
            bool result = false;
            driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromSeconds(2);
            try {
                element.Click();
                result = true;
                ErrorsChecker(driver);
                driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromMinutes(2);
            }
            catch (Exception ex) {
                if (ex.Message.Contains("is not clickable at point") ||
                    ex.Message.Contains("Exception has been thrown by the target of an invocation")) {
                    result = false;
                    ErrorsChecker(driver);
                    driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromMinutes(2);
                }
                else {
                    ErrorsChecker(driver);
                    Assert.Fail(ex.Message);
                }
            }

            driver.Manage().Timeouts().ImplicitWait = TimeSpan.FromMinutes(2);
            return result;
        }

        public static void WaitForPageLoaded(IWebDriver driver) {
            object pageState = new WebDriverWait(driver, TimeSpan.FromSeconds(60)).Until(d =>
                ((IJavaScriptExecutor) d).ExecuteScript("return document.readyState"));
            while (!pageState.Equals("complete")) {
                pageState = new WebDriverWait(driver, TimeSpan.FromSeconds(10)).Until(d =>
                    ((IJavaScriptExecutor) d).ExecuteScript("return document.readyState"));
            }

            Thread.Sleep(1000);
        }

        public static string StartRecordVideo() {
            _videorec = new ScreenCaptureJob {
                CaptureRectangle = new Rectangle(0, 0, 1920, 1040),
                ScreenCaptureVideoProfile = {Quality = 20},
                ShowFlashingBoundary = false
            };

            string path = Assembly.GetCallingAssembly().CodeBase;
            string finalPath = path.Substring(0, path.LastIndexOf("bin", StringComparison.Ordinal)) + "ErrorsVideos\\" +
                               TestContext.CurrentContext.Test.FullName + "_" +
                               DateTime.Now.ToString("yyyy-MM-ddTHH:mm:ss").Replace(':', '_') +
                               ".wmv";
            finalPath = finalPath.Replace("file:///", "").Replace("/", "\\");

            _videorec.OutputScreenCaptureFileName = finalPath;
            _videorec.Start();

            return finalPath;
        }

        private static bool isPageLoadCompleted(IWebDriver driver) {
            ReadOnlyCollection<IWebElement> loadWidgets = driver.FindElements(By.XPath("//*[@class='widget-loading']"));
            if (loadWidgets.Count > 0) {
                Thread.Sleep(500);
                try {
                    foreach (IWebElement loadWidget in loadWidgets) {
                        if (IsElementDisplayed(driver, loadWidget)) {
                            return false;
                        }
                    }
                }
                catch {
                    return false;
                }
            }

            return true;
        }

        public static string GetEnvironment(string outsideUrl = null) {
            string url = outsideUrl.IsNullOrEmpty() ? ConfigurationManager.AppSettings["url"] : outsideUrl;

            if (url.Contains(".uat.") && !url.Contains("stg.uat.")) {
                return "UAT";
            }

            if (url.Contains(".rc.")) {
                return "RC";
            }

            if (url.Contains("stg.uat")) {
                return "STG UAT";
            }

            if (url.Contains("prod.com")) {
                return "PROD";
            }

            return "UNKNOWN";
        }

        public static void UpdateTramServer(bool setCorrectTramsServer) {
            //Common.LoadConfiguration("login");

            string ip = null;
            string username = null;
            string password = null;

            LoadConfiguration("server_connections");

            if (ConfigurationManager.AppSettings["url"].Contains("rc.te.tld")) {
                ip = ModuleData["rc_server_name"];
                username = ModuleData["rc_username"];
                password = ModuleData["rc_password"];
            }
            else if (ConfigurationManager.AppSettings["url"].Contains("uat.te.tld")) {
                ip = ModuleData["uat_server_name"];
                username = ModuleData["uat_username"];
                password = ModuleData["uat_password"];
            }

            string correctTrams = "0.0.0.0:24002/TRAMSService.svc/v2";
            string incorrectTrams = "0.0.0.0:24009/TRAMSService.svc/v2";

            string path =
                Path.Combine(
                    Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location) ??
                    throw new InvalidOperationException(), "Resources");

            string fullPath = ConfigurationManager.AppSettings["url"].Contains("stg.uat.te.tld")
                ? "C:\\TravelEdge\\STG\\Applications\\Api\\Web.config"
                : "C:\\TravelEdge\\Applications\\Api\\web.config";

            string cmdCommand;
            if (setCorrectTramsServer) {
                cmdCommand =
                    $"{path}\\PsExec.exe \\\\{ip} -accepteula -u \\{username} -p {password} -s powershell -Command " +
                    $"\"(gc {fullPath}) -replace '{incorrectTrams}', '{correctTrams}' | Out-File {fullPath} -Encoding 'UTF8'\"";

                Logger.Info($"{TestContext.CurrentContext.Test.FullName} - Trams configured to be UP");
            }
            else {
                cmdCommand =
                    $"{path}\\PsExec.exe \\\\{ip} -accepteula -u \\{username} -p {password} -s powershell -Command " +
                    $"\"(gc {fullPath}) -replace '{correctTrams}', '{incorrectTrams}' | Out-File {fullPath} -Encoding 'UTF8'\"";

                Logger.Info($"{TestContext.CurrentContext.Test.FullName} - Trams configured to be DOWN");
            }

            CommandLine(cmdCommand);
        }

        public static void CommandLine(string command) {
            System.Diagnostics.Process process = new System.Diagnostics.Process();
            System.Diagnostics.ProcessStartInfo startInfo = new System.Diagnostics.ProcessStartInfo {
                WindowStyle = System.Diagnostics.ProcessWindowStyle.Hidden, //Run hidden window
                FileName = "cmd.exe",
                Arguments = "/C " + command,
                //Verb = "runas" //Run as Administrator
            };
            process.StartInfo = startInfo;
            process.Start();
            Thread.Sleep(10000);
        }

        public static string GetClient(IWebDriver driver) {
            string cookie = driver.Manage().Cookies.GetCookieNamed("user").Value;
            string[] cookies = cookie.Split('&');
            string agentId = null;
            foreach (var cook in cookies) {
                string[] cookS = cook.Split('=');
                if (cookS[0] == "AgentId") {
                    agentId = cookS[1];
                    break;
                }
            }

            Dictionary<string, string> data = GetDataFromDbDictionary(
                "SELECT TOP 1 c.ProfileName, COUNT(cc.ClientId) Ccount " +
                "FROM ClientCompanions cc " +
                "JOIN Clients c " +
                "ON c.ClientId = cc.ClientId " +
                "WHERE cc.IsActive = 1 " +
                $"AND c.AgentId = {agentId} " +
                "GROUP BY cc.ClientId, c.ProfileName " +
                "ORDER BY Ccount DESC");

            if (data.Count > 0) {
                foreach (string dataKey in data.Keys) {
                    return dataKey;
                }
            }

            return null;
        }

        public static string EncryptString(string decryptedValue) {
            byte[] initVectorBytes = Encoding.UTF8.GetBytes(InitVector);
            byte[] plainTextBytes = Encoding.UTF8.GetBytes(decryptedValue);
            PasswordDeriveBytes password = new PasswordDeriveBytes(passPhrase, null);
            byte[] keyBytes = password.GetBytes(Keysize / 8);
            RijndaelManaged symmetricKey = new RijndaelManaged {Mode = CipherMode.CBC};
            ICryptoTransform encryptor = symmetricKey.CreateEncryptor(keyBytes, initVectorBytes);
            MemoryStream memoryStream = new MemoryStream();
            CryptoStream cryptoStream = new CryptoStream(memoryStream, encryptor, CryptoStreamMode.Write);
            cryptoStream.Write(plainTextBytes, 0, plainTextBytes.Length);
            cryptoStream.FlushFinalBlock();
            byte[] cipherTextBytes = memoryStream.ToArray();
            memoryStream.Close();
            cryptoStream.Close();
            return Convert.ToBase64String(cipherTextBytes);
        }

        public static string DecryptString(string encryptedValue) {
            byte[] initVectorBytes = Encoding.UTF8.GetBytes(InitVector);
            byte[] cipherTextBytes = Convert.FromBase64String(encryptedValue);
            PasswordDeriveBytes password = new PasswordDeriveBytes(passPhrase, null);
            byte[] keyBytes = password.GetBytes(Keysize / 8);
            RijndaelManaged symmetricKey = new RijndaelManaged {Mode = CipherMode.CBC};
            ICryptoTransform decryptor = symmetricKey.CreateDecryptor(keyBytes, initVectorBytes);
            MemoryStream memoryStream = new MemoryStream(cipherTextBytes);
            CryptoStream cryptoStream = new CryptoStream(memoryStream, decryptor, CryptoStreamMode.Read);
            byte[] plainTextBytes = new byte[cipherTextBytes.Length];
            int decryptedByteCount = cryptoStream.Read(plainTextBytes, 0, plainTextBytes.Length);
            memoryStream.Close();
            cryptoStream.Close();
            return Encoding.UTF8.GetString(plainTextBytes, 0, decryptedByteCount);
        }
    }

    public class ExcelLib
    {
        public static string Path = Assembly.GetCallingAssembly().CodeBase;
        public static string ActualPath = Path.Substring(0, Path.LastIndexOf("bin", StringComparison.Ordinal));
        public static string ProjectPath = new Uri(ActualPath).LocalPath;
        public static string WorkbookPath = ProjectPath + "\\TestData\\TestData.xlsx";
        private Microsoft.Office.Interop.Excel.Application excelApp;
        private Microsoft.Office.Interop.Excel.Workbook excelWorkbook;
        private Microsoft.Office.Interop.Excel.Sheets excelSheets;
        private Microsoft.Office.Interop.Excel.Worksheet excelWorksheet;
        public Microsoft.Office.Interop.Excel.Range XlRange;

        /// <summary>
        /// Fetching the excel sheet data
        /// </summary>
        /// <param name="sheetName"></param>
        /// <param name="cellAddress"></param>
        /// <returns></returns>

        public String GetExcelData(String sheetName, String cellAddress) {
            excelApp = new Microsoft.Office.Interop.Excel.Application();
            excelWorkbook = excelApp.Workbooks.Open(WorkbookPath);

            excelSheets = excelWorkbook.Sheets;

            excelWorksheet = (Microsoft.Office.Interop.Excel.Worksheet) excelSheets.Item[sheetName];
            XlRange = excelWorksheet.UsedRange;

            //going for map
            //Excel.Range usedRange = excelWorksheet.UsedRange;
            //Excel.Range lastCell = usedRange.SpecialCells(Excel.XlCellType.xlCellTypeLastCell, Type.Missing);
            //int lastRowNum = lastCell.Row;
            //for(int r=1;r<= lastRowNum; r++)
            //{
            //    for (int c = 1; c <= 2; c++)
            //    {
            //        Excel.Range cell = usedRange.Cells[r, c];//as Excel.Range;
            //        string s=cell.Value.ToString();
            //        //Excel.Range dataBinder1 = (Excel.Range)excelWorksheet.Cells[r, c];
            //       // string myData1 = dataBinder1.Value.ToString();
            //    }
            //}

            Microsoft.Office.Interop.Excel.Range dataBinder = excelWorksheet.Range[cellAddress, cellAddress];

            // Excel.Range dataBinder1 = (Excel.Range)excelWorksheet.Cells[lastCell.Row + 1, 3]; //row // column

            string myData = dataBinder.Value.ToString();
            excelWorkbook.Close();

            excelApp.Quit();

            return myData;
        }
    }
}