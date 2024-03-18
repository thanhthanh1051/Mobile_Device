using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using System.Threading;
using System.Net.Configuration;
using System.Collections.Generic;
using Microsoft.Office.Interop.Excel;
using OpenQA.Selenium.Support.UI;
using System.Net.NetworkInformation;
using System.Runtime.InteropServices;
using OfficeOpenXml;
using System.IO;
using Newtonsoft.Json;
using SeleniumExtras.WaitHelpers;

namespace do_an.CLIENT
{
    [TestClass]
    public class Favorite
    {
        IWebDriver driver = new ChromeDriver();
        [TestInitialize]
        public void Init()
        {
            driver.Manage().Window.Maximize();
            Thread.Sleep(1000);

            driver.Url = "http://localhost:81/";
            driver.Navigate();
            Thread.Sleep(1000);
            driver.Manage().Window.Maximize();
            Thread.Sleep(1000);
        }
        public void Login(string email, string password)
        {
            bool status = true;
            try
            {
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
                        enterEmail.SendKeys(email);
                    }
                    Thread.Sleep(2000);
                    var enterPassword = driver.FindElement(By.XPath("/html[1]/body[1]/div[5]/div[1]/div[1]/div[1]/div[1]/div[2]/form[1]/input[2]"));
                    Thread.Sleep(1000);
                    status = enterPassword != null;
                    if (status)
                    {
                        enterPassword.SendKeys(password);
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
            //Assert.IsTrue(status);
        }
        public void Logout()
        {
            driver.FindElement(By.XPath("/html/body/div[1]/header/div/div[1]/div/div[3]/div/a[1]/div/span[1]")).Click();
            driver.FindElement(By.XPath("/html/body/div[1]/div[2]/div/div/div[2]/div[1]/div[2]/form/div/a")).Click();
            Thread.Sleep(3000);
        }
        [TestMethod]
        [DataTestMethod]
        [DynamicData(nameof(GetFavoriteCredentialsFromExcel), DynamicDataSourceType.Method)]
        public void AddFavorite(string email, string password)
        {
            bool status = true;
            try
            {
                Login(email, password);
                Thread.Sleep(2000);
                status = driver != null;
                if (status)
                {
                    Thread.Sleep(2000);
                    var btnBuy = driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/div[2]/div[2]/div[2]/div[1]/div[1]/a[1]/div[3]/div[1]/h5[1]/span[1]"));
                    status = btnBuy != null;
                    if (status)
                    {
                        btnBuy.Click();
                    }
                    Thread.Sleep(1000);
                    var enterPressfav = driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/div[2]/div[1]/div[1]/div[1]/div[1]/div[2]/div[2]/a[2]"));
                    Thread.Sleep(1000);
                    status = enterPressfav != null;
                    if (status)
                    {
                        enterPressfav.Click();
                    }
                    try
                    {
                        WebDriverWait wait = new WebDriverWait(driver, TimeSpan.FromSeconds(10));
                        wait.Until(ExpectedConditions.ElementExists(By.XPath("/html/body/div[2]/div")));
                        status = true;
                    }
                    catch (NoSuchElementException)
                    {
                        status = true;
                    }
                }
            }
            catch (Exception ex)
            {
                Assert.IsFalse(status);
                driver.Quit();
            }
            Assert.IsTrue(status);
        }
        [TestMethod]
        [DataTestMethod]
        [DynamicData(nameof(GetFavoriteCredentialsFromExcel), DynamicDataSourceType.Method)]
        public void Test_removeFav(string email, string password, string numberExpect)
        {
            string actual_result = "";
            string result = "";
            bool status = true;
            try
            {
                Login(email, password);
                Thread.Sleep(2000);
                status = driver != null;
                if (status)
                {
                    var btnFav = driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/header[1]/div[1]/div[1]/div[1]/div[3]/div[1]/a[2]/div[1]/span[1]"));
                    status = btnFav != null;
                    if (status)
                    {
                        btnFav.Click();
                    }
                    Thread.Sleep(1000);
                    var enterDeletefav = driver.FindElement(By.XPath("//span[normalize-space()='delete']"));
                    status = enterDeletefav != null;
                    if (status)
                    {
                        enterDeletefav.Click();
                    }
                    Thread.Sleep(1000);
                    var number = driver.FindElement(By.Id("icon-amount-favorite"));
                    status = number != null;
                    if (status)
                    {
                        string textContent = number.Text;
                        if (textContent == numberExpect)
                        {
                            actual_result = textContent;
                            result = "Pass";
                            status = true;
                        }
                        else
                        {
                            actual_result = textContent;
                            result = "Faild";
                            status = false;
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Assert.IsFalse(status);
                driver.Quit();
            }
            Assert.IsTrue(status);
            UpdateFavoriteExcelResult(actual_result, result);
        }
        [TestMethod]
        [DataTestMethod]
        [DynamicData(nameof(GetCommentCredentialsFromExcel), DynamicDataSourceType.Method)]
        public void Test_creatComment(string email, string password, string comment, string expect)
        {
            string actual_result = "";
            string result = "";
            bool status = true;
            try
            {
                Login(email, password);
                Thread.Sleep(2000);
                status = driver != null;
                if (status)
                {
                    Thread.Sleep(2000);
                    var btnBuy = driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/div[2]/div[2]/div[2]/div[1]/div[1]/a[1]/div[3]/div[1]/h5[1]/span[1]"));
                    status = btnBuy != null;
                    if (status)
                    {
                        btnBuy.Click();
                    }
                    Thread.Sleep(1000);
                    var addText = driver.FindElement(By.XPath("//input[@placeholder='What are you thinking?']"));
                    status = addText != null;
                    if (status)
                    {
                        addText.SendKeys(comment);
                    }
                    Thread.Sleep(1000);
                    var btnSave = driver.FindElement(By.XPath("//button[normalize-space()='Share']"));
                    status = btnSave != null;
                    if (status)
                    {
                        btnSave.Click();
                    }
                    IWebElement pElement = driver.FindElement(By.CssSelector("div.panel-body div.media-block div.media-body p.ml-2"));
                    string textContent = pElement.Text;
                    if (textContent == expect)
                    {
                        actual_result = textContent;
                        result = "Pass";
                        status = true;
                    }
                    else
                    {
                        actual_result = textContent;
                        result = "Faild";
                        status = false;
                    }
                }
            }
            catch (Exception ex)
            {
                Assert.IsFalse(status);
                driver.Quit();
            }
            Assert.IsTrue(status);
            UpdateCommentExcelResult(actual_result, result);
        }
        private static IEnumerable<object[]> GetFavoriteCredentialsFromExcel()
        {
            string filePath = @"D:\Baodamchatluong_TH\DO_AN\signUp.xlsx";
            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                var worksheet = package.Workbook.Worksheets[0];
                int rowCount = worksheet.Dimension.Rows;

                for (int row = 2; row <= rowCount; row++)
                {
                    string email = worksheet.Cells[row, 1].Value.ToString();
                    string password = worksheet.Cells[row, 2].Value.ToString();
                    string numberExpect = worksheet.Cells[row, 3].Value.ToString();
                    yield return new string[] { email, password, numberExpect };
                }
            }
        }
        private void UpdateFavoriteExcelResult(string actual_result, string result)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            string filePath = @"D:\Baodamchatluong_TH\DO_AN\signUp.xlsx";
            FileInfo file = new FileInfo(filePath);

            using (ExcelPackage package = new ExcelPackage(file))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0];

                int rowCount = worksheet.Dimension.Rows;

                for (int row = 2; row <= rowCount; row++)
                {
                    worksheet.Cells[row, 4].Value = actual_result;
                    worksheet.Cells[row, 5].Value = result;
                }
                package.Save();
            }
        }

        private static IEnumerable<object[]> GetCommentCredentialsFromExcel()
        {
            string filePath = @"D:\Baodamchatluong_TH\DO_AN\getComment.xlsx";
            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                var worksheet = package.Workbook.Worksheets[0];
                int rowCount = worksheet.Dimension.Rows;

                for (int row = 2; row <= rowCount; row++)
                {
                    string email = worksheet.Cells[row, 1].Value.ToString();
                    string password = worksheet.Cells[row, 2].Value.ToString();
                    string comment = worksheet.Cells[row, 3].Value.ToString();
                    string expect = worksheet.Cells[row, 4].Value.ToString();
                    yield return new string[] { email, password, comment, expect };
                }
            }
        }
        private void UpdateCommentExcelResult(string actual_result, string result)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            string filePath = @"D:\Baodamchatluong_TH\DO_AN\getComment.xlsx";
            FileInfo file = new FileInfo(filePath);

            using (ExcelPackage package = new ExcelPackage(file))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[0];

                int rowCount = worksheet.Dimension.Rows;

                for (int row = 2; row <= rowCount; row++)
                {
                    worksheet.Cells[row, 5].Value = actual_result;
                    worksheet.Cells[row, 6].Value = result;
                }
                package.Save();
            }
        }
    }
}
