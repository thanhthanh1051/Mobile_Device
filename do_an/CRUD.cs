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
namespace do_an
{
    [TestClass]
    public class CRUD
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
                driver.Quit();
            }
            Assert.IsTrue(status);
        }

        //[TestMethod]
        //public void TestMethod1()
        //{
        //    bool status = true;
        //    try
        //    {
        //        driver.Manage().Window.Maximize();
        //        Thread.Sleep(1000);

        //        driver.Url = "http://localhost:81/";
        //        driver.Navigate();
        //        status = driver != null;
        //        if (status)
        //        {
        //            Thread.Sleep(2000);
        //            var iconLogin = driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/header[1]/div[1]/div[1]/div[1]/div[3]/div[1]/a[1]/div[1]/span[1]"));
        //            Thread.Sleep(1000);
        //            status = iconLogin != null;
        //            if (status)
        //            {
        //                iconLogin.Click();
        //            }
        //            Thread.Sleep(1000);
        //            var enterEmail = driver.FindElement(By.XPath("/html[1]/body[1]/div[5]/div[1]/div[1]/div[1]/div[1]/div[2]/form[1]/input[1]"));
        //            Thread.Sleep(1000);
        //            status = enterEmail != null;
        //            if (status)
        //            {
        //                enterEmail.SendKeys("admin123@gmail.com");
        //            }
        //            Thread.Sleep(2000);
        //            var enterPassword = driver.FindElement(By.XPath("/html[1]/body[1]/div[5]/div[1]/div[1]/div[1]/div[1]/div[2]/form[1]/input[2]"));
        //            Thread.Sleep(1000);
        //            status = enterPassword != null;
        //            if (status)
        //            {
        //                enterPassword.SendKeys("admin123");
        //            }
        //            Thread.Sleep(2000);
        //            var clickLogin = driver.FindElement(By.XPath("/html[1]/body[1]/div[5]/div[1]/div[1]/div[1]/div[1]/div[2]/form[1]/button[1]"));
        //            Thread.Sleep(1000);
        //            status = clickLogin != null;
        //            if (status)
        //            {
        //                clickLogin.Click();
        //            }
        //            Thread.Sleep(2000);
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        driver.Quit();
        //    }
        //}

        [TestMethod]
        public void CreateProduct()
        {
             bool status = true;
            try
            {
                driver.SwitchTo().NewWindow(WindowType.Tab);
                driver.Navigate().GoToUrl("http://localhost:81/admin");
                status = driver != null;
                if (status)
                {
                    var createProduct = driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/ul[1]/li[2]/a[1]/span[1]"));
                    Thread.Sleep(1000);
                    status = createProduct != null;
                    if (status)
                    {
                        createProduct.Click();
                    }
                    Thread.Sleep(2000);
                    var clickEnterNameProduct = driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/ul[1]/li[2]/div[1]/div[1]/a[2]"));
                    Thread.Sleep(1000);
                    status = clickEnterNameProduct != null;
                    if (status)
                    {
                        clickEnterNameProduct.Click();
                    }
                    Thread.Sleep(2000);
                    var enterCodeProduct = driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/div[1]/div[1]/div[1]/div[2]/div[1]/form[1]/div[1]/div[1]/input[1]"));
                    Thread.Sleep(1000);
                    status = enterCodeProduct != null;
                    if (status)
                    {
                        enterCodeProduct.SendKeys("ltMac1");
                    }
                    Thread.Sleep(2000);
                    var enterNameProduct = driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/div[1]/div[1]/div[1]/div[2]/div[1]/form[1]/div[1]/div[3]/input[1]"));
                    Thread.Sleep(1000);
                    status = enterNameProduct != null;
                    if (status)
                    {
                        enterNameProduct.SendKeys("MacBook");
                    }
                    Thread.Sleep(2000);
                    var amount = driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/div[1]/div[1]/div[1]/div[2]/div[1]/form[1]/div[1]/div[5]/input[1]"));
                    Thread.Sleep(1000);
                    status = amount != null;
                    if (status)
                    {
                        amount.SendKeys("20");
                    }
                    Thread.Sleep(2000);
                    var image = driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/div[1]/div[1]/div[1]/div[2]/div[1]/form[1]/div[1]/div[7]/input[1]"));
                    Thread.Sleep(1000);
                    status = image != null;
                    if (status)
                    {
                        image.SendKeys("C:\\Users\\NITRO\\Downloads\\iphone-14_1.webp");
                    }
                    Thread.Sleep(2000);
                    var selectCate = driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/div[1]/div[1]/div[1]/div[2]/div[1]/form[1]/div[1]/div[2]/select[1]"));
                    Thread.Sleep(1000);
                    status = selectCate != null;
                    if (status)
                    {
                        var selectElementCategory = new SelectElement(selectCate);
                        selectElementCategory.SelectByValue("1");
                        selectCate.Click();
                    }
                    Thread.Sleep(2000);
                    var selectBrand = driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/div[1]/div[1]/div[1]/div[2]/div[1]/form[1]/div[1]/div[4]/select[1]"));
                    Thread.Sleep(1000);
                    status = selectBrand != null;
                    if (status)
                    {
                        var selectElementBrand = new SelectElement(selectBrand);
                        selectElementBrand.SelectByValue("1");
                        selectBrand.Click();
                    }
                    Thread.Sleep(2000);
                    var priceSell = driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/div[1]/div[1]/div[1]/div[2]/div[1]/form[1]/div[1]/div[6]/input[1]"));
                    Thread.Sleep(1000);
                    status = priceSell != null;
                    if (status)
                    {
                        priceSell.SendKeys("1000");
                    }
                    Thread.Sleep(2000);
                    var priceBuy = driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/div[1]/div[1]/div[1]/div[2]/div[1]/form[1]/div[1]/div[8]/input[1]"));
                    Thread.Sleep(1000);
                    status = priceBuy != null;
                    if (status)
                    {
                        priceBuy.SendKeys("1200");
                    }
                    Thread.Sleep(2000);
                    var storage = driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/div[1]/div[1]/div[1]/div[2]/div[1]/form[1]/div[1]/div[10]/input[1]"));
                    Thread.Sleep(1000);
                    status = storage != null;
                    if (status)
                    {
                        storage.SendKeys("528");
                    }
                    Thread.Sleep(2000);
                    var create = driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/div[1]/div[1]/div[1]/div[2]/div[1]/form[1]/button[1]"));
                    Thread.Sleep(1000);
                    status = create != null;
                    if (status)
                    {
                        create.Click();
                    }
                    Thread.Sleep(2000);
                }
                Thread.Sleep(2000);
            }
            catch (Exception ex)
            {
                driver.Quit();
            }

            try
            {
                if (status)
                {
                    var search = driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/div[1]/div[1]/div[1]/div[2]/div[2]/div[1]/div[1]/div[1]/div[2]/div[1]/label[1]/input[1]"));
                    Thread.Sleep(1000);
                    status = search != null;
                    if (status)
                    {
                        search.SendKeys("ltMac1");
                        Thread.Sleep(1000);
                        var dataempty = driver.FindElement(By.ClassName("dataTables_empty"));
                        Thread.Sleep(1000);
                        status = dataempty != null;
                        if (status)
                        {
                            driver.Quit();
                        }
                        Thread.Sleep(1000);
                        var searchCode = driver.FindElement(By.ClassName("sorting_1"));
                        bool check = (searchCode.Text == "ltMac1");
                        if (!check)
                        {
                            driver.Quit();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                driver.Quit();
            }
            Assert.IsTrue(status);
        }

        [TestMethod]
        public void UpdateProduct()
        {
            bool status = true;
            try
            {
                driver.SwitchTo().NewWindow(WindowType.Tab);
                driver.Navigate().GoToUrl("http://localhost:81/admin");
                status = driver != null;
                if (status)
                {
                    var updateElement = driver.FindElement(By.CssSelector(".nav-item:nth-child(6) span"));
                    Thread.Sleep(1000);
                    status = updateElement != null;
                    if (status)
                    {
                        updateElement.Click();
                        Thread.Sleep(1000);
                        var showList = driver.FindElement(By.LinkText("Show List"));
                        status = showList != null;
                        if (status)
                        {
                            showList.Click();
                        }
                        Thread.Sleep(1000);
                        var update = driver.FindElement(By.LinkText("Sửa"));
                        status = update != null;
                        if (status)
                        {
                            update.Click();
                        }
                        Thread.Sleep(1000);
                        var clickCode = driver.FindElement(By.Name("code"));
                        status = clickCode != null;
                        if(status)
                        {
                            
                            clickCode.SendKeys("ip11");
                        }
                        Thread.Sleep(1000);
                        var NameProduct = driver.FindElement(By.Name("name"));
                        status = NameProduct != null;
                        if (status)
                        {
                            NameProduct.Click();
                            NameProduct.SendKeys("Iphone11");
                        }
                        Thread.Sleep(1000);
                        var amountPro = driver.FindElement(By.Name("amount"));
                        status = amountPro != null;
                        if (status)
                        {

                            amountPro.Click();
                            amountPro.SendKeys("10");
                        }
                        Thread.Sleep(1000);
                        var image = driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/div[1]/div[1]/div[1]/div[2]/div[1]/form[1]/div[1]/div[7]/input[1]"));
                        status = image != null;
                        if (status)
                        {
                            image.SendKeys("C:\\Users\\NITRO\\Downloads\\iphone-14_1.webp");
                        }
                        Thread.Sleep(1000);
                        var color = driver.FindElement(By.Name("color"));
                        status = color != null;
                        if (status)
                        {
                            color.Click();
                            color.SendKeys("#a22525");
                        }
                        Thread.Sleep(1000);
                        var up = driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/div[1]/div[1]/div[1]/div[2]/div[1]/form[1]/button[1]"));
                        status = up != null;
                        if (status)
                        {
                            up.Click();
                        }
                        Thread.Sleep(2000);
                    }
                    else
                    {
                        driver.Quit();
                    }
                }
            }
            catch (Exception ex)
            {
                driver.Quit();
            }
            Assert.IsTrue(status);
        }

        [TestCleanup]
        public void Test_Close()
        {
            driver.Close();
            driver.Quit();
        }
    }
}
