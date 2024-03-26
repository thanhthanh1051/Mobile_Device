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
using OfficeOpenXml;
using System.IO;
namespace do_an.CRUD_test
{
    [TestClass]
    public class product
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
        }
        public void Logout()
        {
            driver.FindElement(By.XPath("/html/body/div[1]/header/div/div[1]/div/div[3]/div/a[1]/div/span[1]")).Click();
            driver.FindElement(By.XPath("/html/body/div[1]/div[2]/div/div/div[2]/div[1]/div[2]/form/div/a")).Click();
            Thread.Sleep(3000);
        }

        [TestMethod]
        [DataTestMethod]
        //[DataRow(CRUD_data.Product.createProduct.Consts.code,
        //    CRUD_data.Product.createProduct.Consts.name,
        //    CRUD_data.Product.createProduct.Consts.amount,
        //    CRUD_data.Product.createProduct.Consts.image,
        //    CRUD_data.Product.createProduct.Consts.category,
        //    CRUD_data.Product.createProduct.Consts.brand,
        //    CRUD_data.Product.createProduct.Consts.priceSell,
        //    CRUD_data.Product.createProduct.Consts.priceBuy,
        //    CRUD_data.Product.createProduct.Consts.storage
        //    )]
        [DynamicData(nameof(GetAddProductCredentialsFromExcel), DynamicDataSourceType.Method)]
        public void CreateProduct(string email, string password, string code, string name, string amount, string image, string color, string category, string brand, string priceSell, string priceBuy, string storage, string rowValue)
        {
            int row = int.Parse(rowValue);
            string actual_result = "";
            Login(email, password);
            Thread.Sleep(1000);
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
                        enterCodeProduct.SendKeys(code);
                    }
                    Thread.Sleep(2000);
                    var enterNameProduct = driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/div[1]/div[1]/div[1]/div[2]/div[1]/form[1]/div[1]/div[3]/input[1]"));
                    Thread.Sleep(1000);
                    status = enterNameProduct != null;
                    if (status)
                    {
                        enterNameProduct.SendKeys(name);
                    }
                    Thread.Sleep(2000);
                    var addamount = driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/div[1]/div[1]/div[1]/div[2]/div[1]/form[1]/div[1]/div[5]/input[1]"));
                    Thread.Sleep(1000);
                    status = amount != null;
                    if (status)
                    {
                        addamount.SendKeys(amount);
                    }
                    Thread.Sleep(2000);
                    var addimage = driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/div[1]/div[1]/div[1]/div[2]/div[1]/form[1]/div[1]/div[7]/input[1]"));
                    Thread.Sleep(1000);
                    status = addimage != null;
                    if (status)
                    {
                        try
                        {
                            addimage.SendKeys(image);
                        }
                        catch { }
                    }
                    Thread.Sleep(2000);
                    var addcolor = driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/div[1]/div[1]/div[1]/div[2]/div[1]/form[1]/div[1]/div[9]/input[1]"));
                    Thread.Sleep(1000);
                    status = addcolor != null;
                    if (status)
                    {
                        addcolor.SendKeys(color);
                    }
                    Thread.Sleep(2000);
                    var selectCate = driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/div[1]/div[1]/div[1]/div[2]/div[1]/form[1]/div[1]/div[2]/select[1]"));
                    Thread.Sleep(1000);
                    status = selectCate != null;
                    if (status)
                    {
                        var selectElementCategory = new SelectElement(selectCate);
                        selectElementCategory.SelectByValue(category);
                        selectCate.Click();
                    }
                    Thread.Sleep(2000);
                    var selectBrand = driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/div[1]/div[1]/div[1]/div[2]/div[1]/form[1]/div[1]/div[4]/select[1]"));
                    Thread.Sleep(1000);
                    status = selectBrand != null;
                    if (status)
                    {
                        var selectElementBrand = new SelectElement(selectBrand);
                        selectElementBrand.SelectByValue(brand);
                        selectBrand.Click();
                    }
                    Thread.Sleep(2000);
                    var addpriceSell = driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/div[1]/div[1]/div[1]/div[2]/div[1]/form[1]/div[1]/div[6]/input[1]"));
                    Thread.Sleep(1000);
                    status = priceSell != null;
                    if (status)
                    {
                        addpriceSell.SendKeys(priceSell);
                    }
                    Thread.Sleep(2000);
                    var addpriceBuy = driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/div[1]/div[1]/div[1]/div[2]/div[1]/form[1]/div[1]/div[8]/input[1]"));
                    Thread.Sleep(1000);
                    status = priceBuy != null;
                    if (status)
                    {
                        addpriceBuy.SendKeys(priceBuy);
                    }
                    Thread.Sleep(2000);
                    var addstorage = driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/div[1]/div[1]/div[1]/div[2]/div[1]/form[1]/div[1]/div[10]/input[1]"));
                    Thread.Sleep(1000);
                    status = storage != null;
                    if (status)
                    {
                        addstorage.SendKeys(storage);
                    }
                    Thread.Sleep(2000);
                    var create = driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/div[1]/div[1]/div[1]/div[2]/div[1]/form[1]/button[1]"));
                    Thread.Sleep(1000);
                    status = create != null;
                    if (status)
                    {
                        create.Click();
                    }
                    Thread.Sleep(1000);
                    var searchAgain = driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/div[1]/div[1]/div[1]/div[2]/div[2]/div[1]/div[1]/div[1]/div[2]/div[1]/label[1]/input[1]"));
                    status = searchAgain != null;
                    if (status)
                    {
                            searchAgain.SendKeys(name);
                            Thread.Sleep(1000);
                            try
                            {
                                var searchPro = driver.FindElement(By.ClassName("sorting_1"));
                                status = searchPro != null;
                                if (status)
                                {
                                    actual_result = searchPro.Text;
                                }
                            }
                            catch (Exception ex)
                            {
                                var dataempty = driver.FindElement(By.ClassName("dataTables_empty"));
                                Thread.Sleep(1000);
                                status = dataempty != null;
                                if (status)
                                {
                                    actual_result = dataempty.Text;
                                }
                                Thread.Sleep(1000);
                            }
                    }
                }
                Thread.Sleep(2000);
            }
            catch (Exception ex)
            {
                try
                {
                    string validateCode = driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/div[1]/div[1]/div[1]/div[2]/div[1]/form[1]/div[1]/div[1]/span[1]")).Text;
                    if (validateCode != null)
                    {
                        actual_result = validateCode;
                    }
                    Thread.Sleep(1000);
                }
                catch { }
                try
                {
                    string validateName = driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/div[1]/div[1]/div[1]/div[2]/div[1]/form[1]/div[1]/div[3]/span[1]")).Text;
                    if (validateName != null)
                    {
                        actual_result = actual_result + validateName;
                    }
                    Thread.Sleep(1000);
                }
                catch
                { }
                try
                {
                    string validateAmount = driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/div[1]/div[1]/div[1]/div[2]/div[1]/form[1]/div[1]/div[5]/span[1]")).Text;
                    if (validateAmount != null)
                    {
                        actual_result = actual_result + validateAmount;
                    }
                    Thread.Sleep(1000);
                }
                catch { }
                try
                {
                    string validateImage = driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/div[1]/div[1]/div[1]/div[2]/div[1]/form[1]/div[1]/div[7]/span[1]")).Text;
                    if (validateImage != null)
                    {
                        actual_result = actual_result + validateImage;
                    }
                    Thread.Sleep(1000);
                }
                catch { }
                try
                {
                    string validateCategory = driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/div[1]/div[1]/div[1]/div[2]/div[1]/form[1]/div[1]/div[2]/span[1]")).Text;
                    if (validateCategory != null)
                    {
                        actual_result = actual_result + validateCategory;
                    }
                    Thread.Sleep(1000);
                }
                catch { }
                try
                {
                    string validateBrand = driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/div[1]/div[1]/div[1]/div[2]/div[1]/form[1]/div[1]/div[4]/span[1]")).Text;
                    if (validateBrand != null)
                    {
                        actual_result = actual_result + validateBrand;
                    }
                    Thread.Sleep(1000);
                }
                catch { }
                try
                {
                    string validatePriceBuy = driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/div[1]/div[1]/div[1]/div[2]/div[1]/form[1]/div[1]/div[6]/span[1]")).Text;
                    if (validatePriceBuy != null)
                    {
                        actual_result = actual_result + validatePriceBuy;
                    }
                    Thread.Sleep(1000);
                }
                catch { }
                try
                {
                    string validatePriceSell = driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/div[1]/div[1]/div[1]/div[2]/div[1]/form[1]/div[1]/div[8]/span[1]")).Text;
                    if (validatePriceSell != null)
                    {
                        actual_result = actual_result + validatePriceSell;
                    }
                    Thread.Sleep(1000);
                }
                catch { }
                try
                {
                    string validateStorage = driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/div[1]/div[1]/div[1]/div[2]/div[1]/form[1]/div[1]/div[10]/span[1]")).Text;
                    if (validateStorage != null)
                    {
                        actual_result = actual_result + validateStorage;
                    }
                    Thread.Sleep(1000);
                }
                catch { }
            }
            AddProductExcelResult(actual_result, row);
        }

        [TestMethod]
        [DataTestMethod]
        //[DataRow(CRUD_data.Product.createProduct.Consts.code,
        //    CRUD_data.Product.createProduct.Consts.name,
        //    CRUD_data.Product.createProduct.Consts.amount,
        //    CRUD_data.Product.createProduct.Consts.image,
        //    CRUD_data.Product.createProduct.Consts.category,
        //    CRUD_data.Product.createProduct.Consts.brand,
        //    CRUD_data.Product.createProduct.Consts.priceSell,
        //    CRUD_data.Product.createProduct.Consts.priceBuy,
        //    CRUD_data.Product.createProduct.Consts.storage
        //    )]
        [DynamicData(nameof(GetUpdateProductCredentialsFromExcel), DynamicDataSourceType.Method)]
        public void UpdateProduct(string email, string password, string code, string name, string newName, string amount, string image, string color, string category, string brand, string priceSell, string priceBuy, string storage, string rowValue)
        {
            int row = int.Parse(rowValue);
            string actual_result = "";
            Login(email, password);
            Thread.Sleep(1000);
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
                        var search = driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/div[1]/div[1]/div[1]/div[2]/div[2]/div[1]/div[1]/div[1]/div[2]/div[1]/label[1]/input[1]"));
                        status = search != null;
                        if (status)
                        {
                            search.SendKeys(name);
                        }
                        var update = driver.FindElement(By.LinkText("Sửa"));
                        status = update != null;
                        if (status)
                        {
                            update.Click();
                        }
                        Thread.Sleep(1000);
                        var clickCode = driver.FindElement(By.Name("code"));
                        status = clickCode != null;
                        if (status)
                        {
                            clickCode.Clear();
                            clickCode.SendKeys(code);
                        }
                        Thread.Sleep(1000);
                        var NameProduct = driver.FindElement(By.Name("name"));
                        status = NameProduct != null;
                        if (status)
                        {
                            NameProduct.Clear();
                            NameProduct.SendKeys(newName);
                        }
                        Thread.Sleep(1000);
                        var amountPro = driver.FindElement(By.Name("amount"));
                        status = amountPro != null;
                        if (status)
                        {

                            amountPro.Clear();
                            amountPro.SendKeys(amount);
                        }
                        Thread.Sleep(1000);
                        var addimage = driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/div[1]/div[1]/div[1]/div[2]/div[1]/form[1]/div[1]/div[7]/input[1]"));
                        Thread.Sleep(1000);
                        status = addimage != null;
                        if (status)
                        {
                            try
                            {
                                addimage.SendKeys(image);
                            }
                            catch { }
                        }
                        Thread.Sleep(2000);
                        var addcolor = driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/div[1]/div[1]/div[1]/div[2]/div[1]/form[1]/div[1]/div[9]/input[1]"));
                        Thread.Sleep(1000);
                        status = addcolor != null;
                        if (status)
                        {
                            addcolor.SendKeys(color);
                        }
                        Thread.Sleep(2000);
                        var selectCate = driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/div[1]/div[1]/div[1]/div[2]/div[1]/form[1]/div[1]/div[2]/select[1]"));
                        Thread.Sleep(1000);
                        status = selectCate != null;
                        if (status)
                        {
                            var selectElementCategory = new SelectElement(selectCate);
                            selectElementCategory.SelectByValue(category);
                            selectCate.Click();
                        }
                        Thread.Sleep(2000);
                        var selectBrand = driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/div[1]/div[1]/div[1]/div[2]/div[1]/form[1]/div[1]/div[4]/select[1]"));
                        Thread.Sleep(1000);
                        status = selectBrand != null;
                        if (status)
                        {
                            var selectElementBrand = new SelectElement(selectBrand);
                            selectElementBrand.SelectByValue(brand);
                            selectBrand.Click();
                        }
                        Thread.Sleep(2000);
                        var addpriceBuy = driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/div[1]/div[1]/div[1]/div[2]/div[1]/form[1]/div[1]/div[6]/input[1]"));
                        Thread.Sleep(1000);
                        status = addpriceBuy != null;
                        if (status)
                        {
                            addpriceBuy.Clear();
                            addpriceBuy.SendKeys(priceSell);
                        }
                        Thread.Sleep(2000);
                        var addpriceSell = driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/div[1]/div[1]/div[1]/div[2]/div[1]/form[1]/div[1]/div[8]/input[1]"));
                        Thread.Sleep(1000);
                        status = addpriceSell != null;
                        if (status)
                        {
                            addpriceSell.Clear();
                            addpriceSell.SendKeys(priceBuy);
                        }
                        Thread.Sleep(2000);
                        var addstorage = driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/div[1]/div[1]/div[1]/div[2]/div[1]/form[1]/div[1]/div[10]/input[1]"));
                        Thread.Sleep(1000);
                        status = addstorage != null;
                        if (status)
                        {
                            addstorage.Clear();
                            addstorage.SendKeys(storage);
                        }
                        Thread.Sleep(2000);
                        var create = driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/div[1]/div[1]/div[1]/div[2]/div[1]/form[1]/button[1]"));
                        Thread.Sleep(1000);
                        status = create != null;
                        if (status)
                        {
                            create.Click();
                        }
                        Thread.Sleep(1000);
                        var searchAgain = driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/div[1]/div[1]/div[1]/div[2]/div[2]/div[1]/div[1]/div[1]/div[2]/div[1]/label[1]/input[1]"));
                        status = searchAgain != null;
                        if (status)
                        {
                            searchAgain.SendKeys(newName);
                            Thread.Sleep(1000);
                            try
                            {
                                var searchPro = driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/div[1]/div[1]/div[1]/div[2]/div[2]/div[1]/div[1]/div[2]/div[1]/table[1]/tbody[1]/tr[1]/td[2]"));
                                status = searchPro != null;
                                if (status)
                                {
                                    actual_result = searchPro.Text;
                                }
                            }
                            catch (Exception ex)
                            {
                                var dataempty = driver.FindElement(By.ClassName("dataTables_empty"));
                                Thread.Sleep(1000);
                                status = dataempty != null;
                                if (status)
                                {
                                    actual_result = dataempty.Text;
                                }
                                Thread.Sleep(1000);
                            }
                        }
                    }
                    Thread.Sleep(2000);
                }
            }
            catch (Exception ex)
            {
                try
                {
                    string validateCode = driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/div[1]/div[1]/div[1]/div[2]/div[1]/form[1]/div[1]/div[1]/span[1]")).Text;
                    if (validateCode != null)
                    {
                        actual_result = validateCode;
                    }
                    Thread.Sleep(1000);
                }
                catch { }
                try
                {
                    string validateName = driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/div[1]/div[1]/div[1]/div[2]/div[1]/form[1]/div[1]/div[3]/span[1]")).Text;
                    if (validateName != null)
                    {
                        actual_result = actual_result + validateName;
                    }
                    Thread.Sleep(1000);
                }
                catch
                { }
                try
                {
                    string validateAmount = driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/div[1]/div[1]/div[1]/div[2]/div[1]/form[1]/div[1]/div[5]/span[1]")).Text;
                    if (validateAmount != null)
                    {
                        actual_result = actual_result + validateAmount;
                    }
                    Thread.Sleep(1000);
                }
                catch { }
                try
                {
                    string validateImage = driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/div[1]/div[1]/div[1]/div[2]/div[1]/form[1]/div[1]/div[7]/span[1]")).Text;
                    if (validateImage != null)
                    {
                        actual_result = actual_result + validateImage;
                    }
                    Thread.Sleep(1000);
                }
                catch { }
                try
                {
                    string validateCategory = driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/div[1]/div[1]/div[1]/div[2]/div[1]/form[1]/div[1]/div[2]/span[1]")).Text;
                    if (validateCategory != null)
                    {
                        actual_result = actual_result + validateCategory;
                    }
                    Thread.Sleep(1000);
                }
                catch { }
                try
                {
                    string validateBrand = driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/div[1]/div[1]/div[1]/div[2]/div[1]/form[1]/div[1]/div[4]/span[1]")).Text;
                    if (validateBrand != null)
                    {
                        actual_result = actual_result + validateBrand;
                    }
                    Thread.Sleep(1000);
                }
                catch { }
                try
                {
                    string validatePriceBuy = driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/div[1]/div[1]/div[1]/div[2]/div[1]/form[1]/div[1]/div[6]/span[1]")).Text;
                    if (validatePriceBuy != null)
                    {
                        actual_result = actual_result + validatePriceBuy;
                    }
                    Thread.Sleep(1000);
                }
                catch { }
                try
                {
                    string validatePriceSell = driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/div[1]/div[1]/div[1]/div[2]/div[1]/form[1]/div[1]/div[8]/span[1]")).Text;
                    if (validatePriceSell != null)
                    {
                        actual_result = actual_result + validatePriceSell;
                    }
                    Thread.Sleep(1000);
                }
                catch { }
                try
                {
                    string validateStorage = driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/div[1]/div[1]/div[1]/div[2]/div[1]/form[1]/div[1]/div[10]/span[1]")).Text;
                    if (validateStorage != null)
                    {
                        actual_result = actual_result + validateStorage;
                    }
                    Thread.Sleep(1000);
                }
                catch { }
            }
            UpdateProductExcelResult(actual_result, row);
        }


        [TestMethod]
        [DataTestMethod]
        //[DataRow(CRUD_data.Product.delete.Consts.code)]
        [DynamicData(nameof(GetDeleteProductCredentialsFromExcel), DynamicDataSourceType.Method)]
        public void DeleteProduct(string email, string password, string code, string rowValue)
        {
            int row = int.Parse(rowValue);
            string actual_result = "";
            Login(email, password);
            Thread.Sleep(1000);
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
                    updateElement.Click();
                    Thread.Sleep(1000);
                    var showList = driver.FindElement(By.LinkText("Show List"));
                    status = showList != null;
                    if (status)
                    {
                        showList.Click();
                    }
                    Thread.Sleep(2000);
                   
                    var search = driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/div[1]/div[1]/div[1]/div[2]/div[2]/div[1]/div[1]/div[1]/div[2]/div[1]/label[1]/input[1]"));
                    status = search != null;
                    if (status)
                    {
                        search.SendKeys(code);
                    }
                    Thread.Sleep(1000);
                    var delete = driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/div[1]/div[1]/div[1]/div[2]/div[2]/div[1]/div[1]/div[2]/div[1]/table[1]/tbody[1]/tr[1]/td[11]/a[1]"));
                    status = delete != null;
                    if (status)
                    {
                        delete.Click();
                    }
                    Thread.Sleep(1000);
                    var checkSearch = driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/div[1]/div[1]/div[1]/div[2]/div[2]/div[1]/div[1]/div[1]/div[2]/div[1]/label[1]/input[1]"));
                    status = checkSearch != null;
                    if (status)
                    {
                        checkSearch.SendKeys(code);
                    }
                    Thread.Sleep(1000);
                    try
                    {
                        var searchPro = driver.FindElement(By.ClassName("sorting_1"));
                        status = searchPro != null;
                        if (status)
                        {
                            actual_result = searchPro.Text;
                        }
                    }
                    catch (Exception ex)
                    {
                        var dataempty = driver.FindElement(By.ClassName("dataTables_empty"));
                        Thread.Sleep(1000);
                        status = dataempty != null;
                        if (status)
                        {
                            actual_result = dataempty.Text;
                        }
                        Thread.Sleep(1000);
                    }
                }
            }
            catch (Exception ex)
            {
                driver.Quit();
            }
            DeleteProductExcelResult(actual_result, row);
        }
        [TestMethod]
        [DataTestMethod]
        [DynamicData(nameof(GetSortProductCredentialsFromExcel), DynamicDataSourceType.Method)]
        public void SortProduct(string email, string password, string sort, string rowValue)
        {
            int row = int.Parse(rowValue);
            string actual_result = "";
            Login(email, password);
            Thread.Sleep(1000);
            bool status = true;
            try
            {
                driver.SwitchTo().NewWindow(WindowType.Tab);
                driver.Navigate().GoToUrl("http://localhost:81/admin");
                status = driver != null;
                if (status)
                {
                    var clickProduct = driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/ul[1]/li[2]/a[1]/span[1]"));
                    status = clickProduct != null;
                    if (status)
                    {
                        clickProduct.Click();
                    }
                    Thread.Sleep(1000);
                    var clickShow = driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/ul[1]/li[2]/div[1]/div[1]/a[1]"));
                    status = clickShow != null;
                    if (status)
                    {
                        clickShow.Click();
                    }
                    Thread.Sleep(1000);
                    if (sort == "Code")
                    {
                        var clickSortName = driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/div[1]/div[1]/div[1]/div[2]/div[2]/div[1]/div[1]/div[2]/div[1]/table[1]/thead[1]/tr[1]/th[1]"));
                        status = clickSortName != null;
                        if (status)
                        {
                            clickSortName.Click();
                        }
                        Thread.Sleep(1000);
                        var itemSort = driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/div[1]/div[1]/div[1]/div[2]/div[2]/div[1]/div[1]/div[2]/div[1]/table[1]/tbody[1]/tr[1]/td[1]"));
                        status = itemSort != null;
                        if (status)
                        {
                            actual_result = itemSort.Text;
                            Thread.Sleep(1000);
                        }
                    }
                    if (sort == "Name")
                    {
                        var clickSortName = driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/div[1]/div[1]/div[1]/div[2]/div[2]/div[1]/div[1]/div[2]/div[1]/table[1]/thead[1]/tr[1]/th[2]"));
                        status = clickSortName != null;
                        if (status)
                        {
                            clickSortName.Click();
                            clickSortName.Click();
                        }
                        Thread.Sleep(1000);
                        var itemSort = driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/div[1]/div[1]/div[1]/div[2]/div[2]/div[1]/div[1]/div[2]/div[1]/table[1]/tbody[1]/tr[1]/td[2]"));
                        status = itemSort != null;
                        if (status)
                        {
                            actual_result = itemSort.Text;
                            Thread.Sleep(1000);
                        }
                    }
                    if (sort == "Category")
                    {
                        var clickSort = driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/div[1]/div[1]/div[1]/div[2]/div[2]/div[1]/div[1]/div[2]/div[1]/table[1]/thead[1]/tr[1]/th[4]"));
                        status = clickSort != null;
                        if (status)
                        {
                            clickSort.Click();
                        }
                        Thread.Sleep(1000);
                        var itemSort = driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/div[1]/div[1]/div[1]/div[2]/div[2]/div[1]/div[1]/div[2]/div[1]/table[1]/tbody[1]/tr[1]/td[4]"));
                        status = itemSort != null;
                        if (status)
                        {
                            actual_result = itemSort.Text;
                            Thread.Sleep(1000);
                        }
                    }
                    if (sort == "Brand")
                    {
                        var clickSort = driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/div[1]/div[1]/div[1]/div[2]/div[2]/div[1]/div[1]/div[2]/div[1]/table[1]/thead[1]/tr[1]/th[5]"));
                        status = clickSort != null;
                        if (status)
                        {
                            clickSort.Click();
                        }
                        Thread.Sleep(1000);
                        var itemSort = driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/div[1]/div[1]/div[1]/div[2]/div[2]/div[1]/div[1]/div[2]/div[1]/table[1]/tbody[1]/tr[1]/td[5]"));
                        status = itemSort != null;
                        if (status)
                        {
                            actual_result = itemSort.Text;
                            Thread.Sleep(1000);
                        }
                    }
                    if (sort == "Amount")
                    {
                        var clickSort = driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/div[1]/div[1]/div[1]/div[2]/div[2]/div[1]/div[1]/div[2]/div[1]/table[1]/thead[1]/tr[1]/th[6]"));
                        status = clickSort != null;
                        if (status)
                        {
                            clickSort.Click();
                            clickSort.Click();
                        }
                        Thread.Sleep(1000);
                        var itemSort = driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/div[1]/div[1]/div[1]/div[2]/div[2]/div[1]/div[1]/div[2]/div[1]/table[1]/tbody[1]/tr[1]/td[6]"));
                        status = itemSort != null;
                        if (status)
                        {
                            actual_result = itemSort.Text;
                            Thread.Sleep(1000);
                        }
                    }
                    if (sort == "Buying")
                    {
                        var clickSort = driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/div[1]/div[1]/div[1]/div[2]/div[2]/div[1]/div[1]/div[2]/div[1]/table[1]/thead[1]/tr[1]/th[7]"));
                        status = clickSort != null;
                        if (status)
                        {
                            clickSort.Click();
                            clickSort.Click();
                        }
                        Thread.Sleep(1000);
                        var itemSort = driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/div[1]/div[1]/div[1]/div[2]/div[2]/div[1]/div[1]/div[2]/div[1]/table[1]/tbody[1]/tr[1]/td[7]"));
                        status = itemSort != null;
                        if (status)
                        {
                            actual_result = itemSort.Text;
                            Thread.Sleep(1000);
                        }
                    }
                    if (sort == "Selling")
                    {
                        var clickSort = driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/div[1]/div[1]/div[1]/div[2]/div[2]/div[1]/div[1]/div[2]/div[1]/table[1]/thead[1]/tr[1]/th[8]"));
                        status = clickSort != null;
                        if (status)
                        {
                            clickSort.Click();
                            clickSort.Click();
                        }
                        Thread.Sleep(1000);
                        var itemSort = driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/div[1]/div[1]/div[1]/div[2]/div[2]/div[1]/div[1]/div[2]/div[1]/table[1]/tbody[1]/tr[1]/td[8]"));
                        status = itemSort != null;
                        if (status)
                        {
                            actual_result = itemSort.Text;
                            Thread.Sleep(1000);
                        }
                    }
                    if (sort == "Created at")
                    {
                        var clickSortCrea = driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/div[1]/div[1]/div[1]/div[2]/div[2]/div[1]/div[1]/div[2]/div[1]/table[1]/thead[1]/tr[1]/th[9]"));
                        status = clickSortCrea != null;
                        if (status)
                        {
                            clickSortCrea.Click();
                        }
                        Thread.Sleep(1000);
                        var itemSort = driver.FindElement(By.XPath("/html[1]/body[1]/div[1]/div[1]/div[1]/div[1]/div[2]/div[2]/div[1]/div[1]/div[2]/div[1]/table[1]/tbody[1]/tr[1]/td[9]"));
                        status = itemSort != null;
                        if (status)
                        {
                            actual_result = itemSort.Text;
                            Thread.Sleep(1000);
                        }
                    }
                }
            }
            catch
            {
                driver.Quit();
            }
            SortProductExcelResult(actual_result, row);
        }
    

        [TestCleanup]
        public void Test_Close()
        {
            //driver.Close();
            driver.Quit();
        }
        private static IEnumerable<object[]> GetAddProductCredentialsFromExcel()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            string filePath = @"D:\Baodamchatluong_TH\DO_AN\TestCaseALL.xlsx";
            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                var worksheet = package.Workbook.Worksheets[9];
                int rowCount = worksheet.Dimension.Rows;

                for (int row = 2; row <= rowCount; row++)
                {
                    string rowValue = worksheet.Cells[row, 1].Value.ToString();
                    string email = worksheet.Cells[row, 3].Value.ToString();
                    string password = worksheet.Cells[row, 4].Value.ToString();
                    object cellValueCode = worksheet.Cells[row, 5].Value;
                    string code = cellValueCode != null ? cellValueCode.ToString() : string.Empty;
                    object cellValueName = worksheet.Cells[row, 6].Value;
                    string name = cellValueName != null ? cellValueName.ToString() : string.Empty;
                    object cellValueImage = worksheet.Cells[row, 7].Value;
                    string image = cellValueImage != null ? cellValueImage.ToString() : string.Empty;
                    object cellValueAmount = worksheet.Cells[row, 8].Value;
                    string amount = cellValueAmount != null ? cellValueAmount.ToString() : string.Empty;
                    object cellValueColor = worksheet.Cells[row, 9].Value;
                    string color = cellValueColor != null ? cellValueColor.ToString() : string.Empty;
                    string category = worksheet.Cells[row, 10].Value.ToString();
                    string brand = worksheet.Cells[row, 11].Value.ToString();
                    object cellValuePriceBuy = worksheet.Cells[row, 12].Value;
                    string price_Buy = cellValuePriceBuy != null ? cellValuePriceBuy.ToString() : string.Empty;
                    object cellValuePriceSell = worksheet.Cells[row, 13].Value;
                    string price_Sell = cellValuePriceSell != null ? cellValuePriceSell.ToString() : string.Empty;
                    object cellValueStorage = worksheet.Cells[row, 14].Value;
                    string storage = cellValueStorage != null ? cellValueStorage.ToString() : string.Empty;
                    yield return new string[] { email, password, code, name, amount, image, color, category, brand, price_Buy, price_Sell, storage , rowValue };
                }
            }
        }
        private void AddProductExcelResult(string actual_result, int row)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            string filePath = @"D:\Baodamchatluong_TH\DO_AN\TestCaseALL.xlsx";
            FileInfo file = new FileInfo(filePath);

            using (ExcelPackage package = new ExcelPackage(file))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[9];
                int rowCount = worksheet.Dimension.Rows;
                string expected = worksheet.Cells[row, 15].Value.ToString();
                if (actual_result == expected)
                {
                    worksheet.Cells[row, 16].Value = actual_result;
                    worksheet.Cells[row, 17].Value = "Pass";
                }
                else
                {
                    worksheet.Cells[row, 16].Value = actual_result;
                    worksheet.Cells[row, 17].Value = "Faild";
                }
                package.Save();
            }
        }
        private static IEnumerable<object[]> GetUpdateProductCredentialsFromExcel()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            string filePath = @"D:\Baodamchatluong_TH\DO_AN\TestCaseALL.xlsx";
            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                var worksheet = package.Workbook.Worksheets[15];
                int rowCount = worksheet.Dimension.Rows;

                for (int row = 2; row <= rowCount; row++)
                {
                    string rowValue = worksheet.Cells[row, 1].Value.ToString();
                    string email = worksheet.Cells[row, 3].Value.ToString();
                    string password = worksheet.Cells[row, 4].Value.ToString();
                    object cellValueCode = worksheet.Cells[row, 5].Value;
                    string code = cellValueCode != null ? cellValueCode.ToString() : string.Empty;
                    object cellValueName = worksheet.Cells[row, 6].Value;
                    string name = cellValueName != null ? cellValueName.ToString() : string.Empty;
                    object cellValueNewName = worksheet.Cells[row, 7].Value;
                    string newname = cellValueNewName != null ? cellValueNewName.ToString() : string.Empty;
                    object cellValueImage = worksheet.Cells[row, 8].Value;
                    string image = cellValueImage != null ? cellValueImage.ToString() : string.Empty;
                    object cellValueAmount = worksheet.Cells[row, 9].Value;
                    string amount = cellValueAmount != null ? cellValueAmount.ToString() : string.Empty;
                    object cellValueColor = worksheet.Cells[row, 10].Value;
                    string color = cellValueAmount != null ? cellValueAmount.ToString() : string.Empty;
                    string category = worksheet.Cells[row, 11].Value.ToString();
                    string brand = worksheet.Cells[row, 12].Value.ToString();
                    object cellValuePriceBuy = worksheet.Cells[row, 13].Value;
                    string price_Buy = cellValuePriceBuy != null ? cellValuePriceBuy.ToString() : string.Empty;
                    object cellValuePriceSell = worksheet.Cells[row, 14].Value;
                    string price_Sell = cellValuePriceSell != null ? cellValuePriceSell.ToString() : string.Empty;
                    object cellValueStorage = worksheet.Cells[row, 15].Value;
                    string storage = cellValueStorage != null ? cellValueStorage.ToString() : string.Empty;
                    yield return new string[] { email, password, code, name, newname, amount, image, color, category, brand, price_Buy, price_Sell, storage, rowValue };
                }
            }
        }
        private void UpdateProductExcelResult(string actual_result, int row)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            string filePath = @"D:\Baodamchatluong_TH\DO_AN\TestCaseALL.xlsx";
            FileInfo file = new FileInfo(filePath);

            using (ExcelPackage package = new ExcelPackage(file))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets[15];
                int rowCount = worksheet.Dimension.Rows;
                string expected = worksheet.Cells[row, 16].Value.ToString();
                if (actual_result == expected)
                {
                    worksheet.Cells[row, 17].Value = actual_result;
                    worksheet.Cells[row, 18].Value = "Pass";
                }
                else
                {
                    worksheet.Cells[row, 17].Value = actual_result;
                    worksheet.Cells[row, 18].Value = "Faild";
                }
                package.Save();
            }
        }
        private static IEnumerable<object[]> GetDeleteProductCredentialsFromExcel()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            string filePath = @"D:\Baodamchatluong_TH\DO_AN\TestCaseALL.xlsx";
            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                var worksheet = package.Workbook.Worksheets["delete product"];
                int rowCount = worksheet.Dimension.Rows;

                for (int row = 2; row <= rowCount; row++)
                {
                    string rowValue = worksheet.Cells[row, 1].Value.ToString();
                    string email = worksheet.Cells[row, 3].Value.ToString();
                    string password = worksheet.Cells[row, 4].Value.ToString();
                    object cellValueCode = worksheet.Cells[row, 5].Value;
                    string code = cellValueCode != null ? cellValueCode.ToString() : string.Empty;
                    yield return new string[] { email, password, code, rowValue };
                }
            }
        }
        private void DeleteProductExcelResult(string actual_result, int row)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            string filePath = @"D:\Baodamchatluong_TH\DO_AN\TestCaseALL.xlsx";
            FileInfo file = new FileInfo(filePath);

            using (ExcelPackage package = new ExcelPackage(file))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets["delete product"];
                int rowCount = worksheet.Dimension.Rows;
                string expected = worksheet.Cells[row, 6].Value.ToString();
                if (actual_result == expected)
                {
                    worksheet.Cells[row, 7].Value = actual_result;
                    worksheet.Cells[row, 8].Value = "Pass";
                }
                else
                {
                    worksheet.Cells[row, 7].Value = actual_result;
                    worksheet.Cells[row, 8].Value = "Faild";
                }
                package.Save();
            }
        }
        private static IEnumerable<object[]> GetSortProductCredentialsFromExcel()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            string filePath = @"D:\Baodamchatluong_TH\DO_AN\TestCaseALL.xlsx";
            using (var package = new ExcelPackage(new FileInfo(filePath)))
            {
                var worksheet = package.Workbook.Worksheets["SortProduct"];
                int rowCount = worksheet.Dimension.Rows;

                for (int row = 2; row <= rowCount; row++)
                {
                    string rowValue = worksheet.Cells[row, 1].Value.ToString();
                    string email = worksheet.Cells[row, 3].Value.ToString();
                    string password = worksheet.Cells[row, 4].Value.ToString();
                    object cellValueSort = worksheet.Cells[row, 5].Value;
                    string sort = cellValueSort != null ? cellValueSort.ToString() : string.Empty;
                    yield return new string[] { email, password, sort, rowValue };
                }
            }
        }
        private void SortProductExcelResult(string actual_result, int row)
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            string filePath = @"D:\Baodamchatluong_TH\DO_AN\TestCaseALL.xlsx";
            FileInfo file = new FileInfo(filePath);

            using (ExcelPackage package = new ExcelPackage(file))
            {
                ExcelWorksheet worksheet = package.Workbook.Worksheets["SortProduct"];
                int rowCount = worksheet.Dimension.Rows;
                string expected = worksheet.Cells[row, 6].Value.ToString();
                if (actual_result == expected)
                {
                    worksheet.Cells[row, 7].Value = actual_result;
                    worksheet.Cells[row, 8].Value = "Pass";
                }
                else
                {
                    worksheet.Cells[row, 7].Value = actual_result;
                    worksheet.Cells[row, 8].Value = "Faild";
                }
                package.Save();
            }
        }
    }
}
