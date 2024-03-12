using Microsoft.VisualStudio.TestTools.UnitTesting;
using OpenQA.Selenium.Chrome;
using OpenQA.Selenium;
using System;
using System.Threading;
using Microsoft.Office.Interop.Excel;
using OpenQA.Selenium.Support.UI;

namespace do_an.CRUD_test
{
    [TestClass]
    public class order
    {
        IWebDriver driver = new ChromeDriver();
        [TestInitialize]
        public void Test_Login()
        {

            bool status = true;
            try
            {
                driver.Manage().Window.Maximize();
                Thread.Sleep(1000);

                driver.Url = "http://localhost:81/";
                driver.Navigate();
                status = driver != null;
                if (status)
                {
                    Thread.Sleep(2000);
                    var iconLogin = driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/header[1]/div[1]/div[1]/div[1]/div[3]/div[1]/a[1]/div[1]/span[1]"));
                    Thread.Sleep(1000);
                    status = iconLogin != null;
                    if (status)
                    {
                        iconLogin.Click();
                    }
                    Thread.Sleep(1000);
                    var enterEmail = driver.FindElement(By.XPath("/html[1]/body[1]/div[5]/div[1]/div[1]/div[1]/div[1]/div[2]/form[1]/input[1]"));
                    Thread.Sleep(1000);
                    status = enterEmail != null;
                    if (status)
                    {
                        enterEmail.SendKeys("admin123@gmail.com");
                    }
                    Thread.Sleep(2000);
                    var enterPassword = driver.FindElement(By.XPath("/html[1]/body[1]/div[5]/div[1]/div[1]/div[1]/div[1]/div[2]/form[1]/input[2]"));
                    Thread.Sleep(1000);
                    status = enterPassword != null;
                    if (status)
                    {
                        enterPassword.SendKeys("admin123");
                    }
                    Thread.Sleep(2000);
                    var clickLogin = driver.FindElement(By.XPath("/html[1]/body[1]/div[5]/div[1]/div[1]/div[1]/div[1]/div[2]/form[1]/button[1]"));
                    Thread.Sleep(1000);
                    status = clickLogin != null;
                    if (status)
                    {
                        clickLogin.Click();
                    }
                    Thread.Sleep(2000);
                }
            }
            catch (Exception ex)
            {
                Assert.IsFalse(status);
                driver.Quit();
            }
            Assert.IsTrue(status);
        }

        public void Cancelled()
        {
            bool status = true;
            try
            {
                driver.SwitchTo().NewWindow(WindowType.Tab);
                driver.Navigate().GoToUrl("http://localhost:81/admin");
                status = driver != null;
                if (status)
                {
                    var clickOrder = driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/ul[1]/li[5]/a[1]/span[1]"));
                    status = clickOrder != null;
                    if (status)
                    {
                        clickOrder.Click();
                    }
                    Thread.Sleep(1000);
                    var clickShow = driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/ul[1]/li[5]/div[1]/div[1]/a[1]"));
                    status = clickShow != null;
                    if (status)
                    {
                        clickShow.Click();
                    }
                    Thread.Sleep(1000);
                    var clickUpdate = driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/div[1]/div[1]/div[1]/div[2]/div[2]/div[1]/div[1]/div[2]/div[1]/table[1]/tbody[1]/tr[1]/td[9]/a[1]"));
                    status = clickUpdate != null;
                    if (status)
                    {
                        clickUpdate.Click();
                    }
                    Thread.Sleep(1000);
                    var clickStatus = driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/div[1]/div[1]/div[1]/div[2]/div[1]/form[1]/div[3]/select[1]"));
                    status = clickStatus != null;
                    if (status)
                    {
                        var selectElement = new SelectElement(clickStatus);
                        selectElement.SelectByValue("2");
                        clickStatus.Click();
                    }
                    Thread.Sleep(1000);
                   
                }
               }
            catch(Exception ex)
            {
                Assert.IsFalse(status);
                driver.Close();
                driver.Quit();
            }
            Assert.IsTrue(status);
            driver.Close();
        }
    }
}
