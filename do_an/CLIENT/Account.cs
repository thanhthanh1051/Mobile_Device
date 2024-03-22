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
    public class Account
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
        [DynamicData(nameof(GetLoginCredentialsFromExcel), DynamicDataSourceType.Method)]
        public void LoginAccount(string email, string password,string expect)
        {
            string actual_result = "";
            string result = "";
            Login(email, password);
            string user_name = driver.FindElement(By.XPath("/html/body/div[1]/header/div/div[2]/div/div/div[3]")).Text;
                if (user_name == expect)
                {
                    actual_result = expect;
                    result = "Pass";
                    Thread.Sleep(2000);
                    Logout();
                }
                else if (string.IsNullOrEmpty(user_name))
                {
                    actual_result = "Đăng nhập thất bại";
                    result = "Faild";
                }
            UpdateExcelResult(actual_result,result);
        }

        [TestMethod]
        [DataTestMethod]
        [DynamicData(nameof(GetSignUpCredentialsFromExcel), DynamicDataSourceType.Method)]
        public void CreateAccount(string name, string email, string password, string confirm_password, string expect)
        {
            string actual_result = "";
            string result = "";
            bool status = true;
            var clickLogin = driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/header[1]/div[1]/div[1]/div[1]/div[3]/div[1]/a[1]/div[1]/span[1]"));
            status = clickLogin != null;
            try
            {
                if(status)
                {
                    clickLogin.Click();
                }
                Thread.Sleep(1000);
                var clickSignUp = driver.FindElement(By.XPath("/html[1]/body[1]/div[5]/div[1]/div[1]/div[1]/div[1]/div[3]/div[1]/div[2]/button[1]"));
                status = clickSignUp != null;
                if (status)
                {
                    clickSignUp.Click();
                }
                Thread.Sleep(1000);
                var enterName = driver.FindElement(By.XPath("/html[1]/body[1]/div[5]/div[1]/div[1]/div[1]/div[1]/div[1]/form[1]/input[1]"));
                status = enterName != null;
                if (status)
                {
                    enterName.SendKeys(name);
                }
                Thread.Sleep(1000);
                var enterEmail = driver.FindElement(By.XPath("/html[1]/body[1]/div[5]/div[1]/div[1]/div[1]/div[1]/div[1]/form[1]/input[2]"));
                status = enterEmail != null;
                if (status)
                {
                    enterEmail.SendKeys(email);
                }
                Thread.Sleep(1000);
                var enterPassword = driver.FindElement(By.XPath("/html[1]/body[1]/div[5]/div[1]/div[1]/div[1]/div[1]/div[1]/form[1]/input[3]"));
                status = enterPassword != null;
                if (status)
                {
                    enterPassword.SendKeys(password);
                }
                Thread.Sleep(1000);
                var enterConfirmPassword = driver.FindElement(By.XPath("/html[1]/body[1]/div[5]/div[1]/div[1]/div[1]/div[1]/div[1]/form[1]/input[4]"));
                status = enterConfirmPassword != null;
                if (status)
                {
                    enterConfirmPassword.SendKeys(confirm_password);
                }
                var signUp = driver.FindElement(By.XPath("/html[1]/body[1]/div[5]/div[1]/div[1]/div[1]/div[1]/div[1]/form[1]/button[1]"));
                status = signUp != null;
                if(status)
                {
                    signUp.Click();
                }
                Thread.Sleep(1000);
                Login(email, password);
                string user_name = driver.FindElement(By.XPath("/html/body/div[1]/header/div/div[2]/div/div/div[3]")).Text;
                if (user_name == expect)
                    {
                       actual_result = expect;
                        result = "Pass";
                       Thread.Sleep(2000);
                       Logout();
                    }
                else
                    {
                       actual_result = "Đăng ký không thành công";
                       result = "Faild";
                    }
            }
            catch (Exception ex)
            {
                driver.Quit();
            }
            UpdateSignUpExcelResult(actual_result,result);
        }
        [TestMethod]
        [DataTestMethod]
        [DynamicData(nameof(GetLogOutCredentialsFromExcel), DynamicDataSourceType.Method)]
        public void Logout(string name, string email, string password,string expect) 
        {
            string actual_result = "";
            string result = "";
            Login(email, password);
            string user_name = driver.FindElement(By.XPath("/html/body/div[1]/header/div/div[2]/div/div/div[3]")).Text;
            if (user_name == name)
            {
                Logout();
                string user_name_logout = driver.FindElement(By.XPath("/html/body/div[1]/header/div/div[2]/div/div/div[3]")).Text;
                if (user_name_logout == expect)
                {
                    actual_result = "Đăng xuất thành công";
                    result = "Pass";
                }
                else
                {
                    actual_result = "Đăng xuất thất bại";
                    result = "Faild";
                }
            }
            else
            {
                actual_result = "Đăng nhập thất bại";
                result = "Faild";
            }
            UpdateLogOutExcelResult(actual_result, result);
        }
        private static IEnumerable<object[]> GetLoginCredentialsFromExcel()
        {
            string filePath = @"D:\Baodamchatluong_TH\DO_AN\loginAccount.xlsx";
            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                var worksheet = package.Workbook.Worksheets[0];
                int rowCount = worksheet.Dimension.Rows;

                for (int row = 2; row <= rowCount; row++)
                {
                    string email = worksheet.Cells[row, 1].Value.ToString();
                    string password = worksheet.Cells[row, 2].Value.ToString();
                    string expect = worksheet.Cells[row, 3].Value.ToString();
                    yield return new string[] { email, password, expect };
                }
            }
        }
        private void UpdateExcelResult(string actual_result,string result)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            string filePath = @"D:\Baodamchatluong_TH\DO_AN\loginAccount.xlsx";
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
        private static IEnumerable<object[]> GetSignUpCredentialsFromExcel()
        {
            string filePath = @"D:\Baodamchatluong_TH\DO_AN\signUp.xlsx";
            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                var worksheet = package.Workbook.Worksheets[0];
                int rowCount = worksheet.Dimension.Rows;

                for (int row = 2; row <= rowCount; row++)
                {
                    string name = worksheet.Cells[row, 1].Value.ToString();
                    string email = worksheet.Cells[row, 2].Value.ToString();
                    string password = worksheet.Cells[row, 3].Value.ToString();
                    string comfirm_password = worksheet.Cells[row, 4].Value.ToString();
                    string expect = worksheet.Cells[row, 5].Value.ToString();
                    yield return new string[] { name, email, password, comfirm_password, expect };
                }
            }
        }
        private void UpdateSignUpExcelResult(string actual_result, string result)
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
                    worksheet.Cells[row, 6].Value = actual_result;
                    worksheet.Cells[row, 7].Value = result;
                }
                package.Save();
            }
        }
        private static IEnumerable<object[]> GetLogOutCredentialsFromExcel()
        {
            string filePath = @"D:\Baodamchatluong_TH\DO_AN\logout.xlsx";
            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                var worksheet = package.Workbook.Worksheets[0];
                int rowCount = worksheet.Dimension.Rows;

                for (int row = 2; row <= rowCount; row++)
                {
                    string name = worksheet.Cells[row, 1].Value.ToString();
                    string email = worksheet.Cells[row, 2].Value.ToString();
                    string password = worksheet.Cells[row, 3].Value.ToString();
                    string expect = worksheet.Cells[row, 4].Value.ToString();
                    yield return new string[] { email, password, expect };
                }
            }
        }
        private void UpdateLogOutExcelResult(string actual_result, string result)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            string filePath = @"D:\Baodamchatluong_TH\DO_AN\logout.xlsx";
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
        [TestCleanup]
        public void clear()
        {
            driver.Quit();
        }
    }
}
