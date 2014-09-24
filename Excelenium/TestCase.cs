/*
 * Excelenium - An excel driven Selenium test suite
 * Copyright (c) Paul Connolly, paul.connolly@managed.info 2014
 * All rights reserved
 */

using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using OpenQA.Selenium;
using OpenQA.Selenium.Firefox;
using NUnit.Framework;

namespace Excelenium
{
    class TestCase
    {
        private String logLevel;
        private String name;
        private String baseurl;
        private List<Operation> operations;
        private IWebDriver driver;

        IWebElement e = null;
        ReadOnlyCollection<IWebElement> eles = null;

        public String getName()
        {
            return name;
        }

        public String getBaseurl()
        {
            return baseurl;
        }

        public TestCase(String logLevel, String name, String baseUrl)
        {
            this.logLevel = logLevel;
            this.name = name;
            this.baseurl = baseUrl;
            this.operations = new List<Operation>();
        }

        public TestCase(String logLevel, System.Array args)
        {
            this.logLevel = logLevel;

            if (args.Length > 0) this.name = (String)args.GetValue(1, 1);
            if (args.Length > 1) this.baseurl = (String)args.GetValue(1, 2);

            this.operations = new List<Operation>();

            if (args.Length > 5) this.operations.Add(new Operation((String)args.GetValue(1, 3), (String)args.GetValue(1, 4), (String)args.GetValue(1, 5), (String)args.GetValue(1, 6), (String)args.GetValue(1, 7)));
        }

        public void addOperation(String action, String method, String what, String actionOnElement, String argument)
        {
            if (action == null && actionOnElement == null)
            {
                LOG.always("Must have either action or actionOnElement to create an operation");
            }
            else
                this.operations.Add(new Operation(action, method, what, actionOnElement, argument));
        }

        public void addOperation(System.Array args)
        {
            if (args.GetValue(1, 1) == null && args.GetValue(1, 4) == null)
            {
                Console.WriteLine("Must have either action or actionOnElement to create an operation");
            }
            else
                this.operations.Add(new Operation(args));
        }

        public bool run()
        {
            bool ret = true;

            try
            {
                driver = new FirefoxDriver();
                // Tell autocomplete to wait - Implicit Wait
                driver.Manage().Timeouts().ImplicitlyWait(TimeSpan.FromSeconds(20));

                for (int i = 0; i < operations.Count; i++)
                {
                    LOG.verbose("Running Test '" + this.name + "' operation " + i);

                    Operation o = operations[i];

                    LOG.verbose("\tBase URL : '" + baseurl + "'");
                    LOG.verbose("\tAction : '" + o.action + "'");
                    LOG.verbose("\tMethod : '" + o.method + "'");
                    LOG.verbose("\tWhat : '" + o.what + "'");
                    LOG.verbose("\tAction On Element : '" + o.actionOnElement + "'");
                    LOG.verbose("\tArgument : '" + o.argument + "'");

                    switch (o.action)
                    {
                        case null:
                            break;
                        case "gotoUrl":
                            driver.Navigate().GoToUrl(baseurl + o.what);
                            e = null;
                            break;

                        case "FindElement":
                            switch (o.method)
                            {
                                case "By.Id":
                                    e = driver.FindElement(By.Id(o.what));
                                    break;

                                case "By.ClassName":
                                    e = driver.FindElement(By.ClassName(o.what));
                                    break;

                                default:
                                    throw new Exception("Unknown method '" + o.method + "' for " + o.action);
                            }
                            break;

                        case "FindSubElement":
                            switch (o.method)
                            {
                                case "By.Id":
                                    e = e.FindElement(By.Id(o.what));
                                    break;

                                case "By.ClassName":
                                    e = e.FindElement(By.ClassName(o.what));
                                    break;

                                default:
                                    throw new Exception("Unknown method '" + o.method + "' for " + o.action);
                            }
                            break;

                        case "SearchElements":
                            // we have a list of elements - search for the one we want
                            foreach (IWebElement ele in eles)
                            {
                                switch (o.method)
                                {
                                    case "Text":
                                        if (ele.Text == o.what)
                                        {
                                            e = ele;
                                        }
                                        break;

                                    default:
                                        throw new Exception("Unknown method '" + o.method + "' for " + o.action);
                                }
                            }

                            break;

                        case "FindElements":
                            switch (o.method)
                            {
                                case "By.Id":
                                    eles = driver.FindElements(By.Id(o.what));
                                    break;

                                case "By.ClassName":
                                    eles = driver.FindElements(By.ClassName(o.what));
                                    break;

                                default:
                                    throw new Exception("Unknown method '" + o.method + "'");
                            }

                            break;

                        case "FindSubElements":
                            switch (o.method)
                            {
                                case "By.Id":
                                    eles = e.FindElements(By.Id(o.what));
                                    break;

                                case "By.ClassName":
                                    eles = e.FindElements(By.ClassName(o.what));
                                    break;

                                default:
                                    throw new Exception("Unknown method '" + o.method + "'");
                            }

                            break;

                        default:
                            throw new Exception("Unknown Action '" + o.action + "'");
                    }

                    if (e != null && o.actionOnElement != null)
                    {
                        switch (o.actionOnElement)
                        {
                            case null:
                                break;
                            case "Clear":
                                e.Clear();
                                break;

                            case "Click":
                                e.Click();
                                System.Threading.Thread.Sleep(5000);
                                e = null;
                                break;

                            case "SendKeys":
                                LOG.full("Sending keys: '" + o.argument + "'");
                                e.SendKeys(o.argument);
                                break;

                            case "Assert.AreEqual":
                                LOG.full("Checking Are Equal: '" + o.argument + "'");
                                Assert.AreEqual(e.Text, o.argument);
                                break;

                            default:
                                throw new Exception("Unknown Action On Element '" + o.actionOnElement + "'");
                        }
                    }
                }
            }
            catch (Exception e)
            {
                LOG.normal(e.Message);
                ret = false;
            }

            finally
            {
                driver.Quit();
            }

            return ret;
        }
    }

    class Operation
    {
        public String action;
        public String method;
        public String what;
        public String actionOnElement;
        public String argument;

        public Operation(String action, String method, String what, String actionOnElement, String argument)
        {
            this.action = action;
            this.method = method;
            this.what = what;
            this.actionOnElement = actionOnElement;
            this.argument = argument;
        }

        public Operation(System.Array args)
        {
            this.action = (String)args.GetValue(1, 1);
            this.method = (String)args.GetValue(1, 2);
            this.what = (String)args.GetValue(1, 3);
            this.actionOnElement = (String)args.GetValue(1, 4); ;
            this.argument = (String)args.GetValue(1, 5); ;
        }
    }
}