using System;
using System.Net;
using AventStack.ExtentReports;
using AventStack.ExtentReports.Reporter;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
namespace Excel
{
    [TestClass]
    public class UnitTest1
    {
        [TestMethod]
        public void TestMethod1()
        {
            IWebDriver driver = new ChromeDriver();


            var reporter = new ExtentHtmlReporter(@".\Report\reports.html");

            var extent = new ExtentReports();

            //real value
            string host = Dns.GetHostName();

            OperatingSystem os = Environment.OSVersion;
            extent.AddSystemInfo("HostName", host);
            extent.AddSystemInfo("operating system", os.ToString());
            extent.AddSystemInfo("Browser", "Google Chrome");

            extent.AttachReporter(reporter);

            var test = extent.CreateTest("excel");


            driver.Navigate().GoToUrl("https://cricket.yahoo.net/");

            driver.Manage().Window.Maximize();


          Screenshot ss=  ((ITakesScreenshot)driver).GetScreenshot();

            ss.SaveAsFile(@".\Screenshots\image.png");

            test.Log(Status.Pass, "Test case pass");

            extent.Flush();
        }
    }
}
