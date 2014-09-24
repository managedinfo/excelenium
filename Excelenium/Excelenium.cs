/*
 * Excelenium - An excel driven Selenium test suite
 * Copyright (c) Paul Connolly, paul.connolly@managed.info 2014
 * All rights reserved
 */

using System;
using System.Diagnostics;
using System.Text;
using Excel = Microsoft.Office.Interop.Excel;

namespace Excelenium
{
    class Excelenium
    {
        static private String inputFile = null;
        static private String testName = null;
        static private String logLevel = "normal";

        static void Main(string[] args)
        {
            int run = 0;
            int skipped = 0;
            int passed = 0;
            int failed = 0;
            int parsed = 0;

            if (!parseArgs(args)) return;

            Excel.Application xlApp;
            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            Excel.Range range;

            xlApp = new Excel.Application();
            xlWorkBook = xlApp.Workbooks.Open(System.IO.Path.GetFullPath(inputFile)); // use get full path to prevent issues with relative paths on the command line
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

            range = xlWorkSheet.UsedRange;

            LOG.verbose("Number of parsed cols: " + range.Columns.Count);
            LOG.verbose("Number of parsed rows: " + range.Rows.Count);

            TestCase tc = null;

            for (int rCnt = 2; rCnt <= range.Rows.Count; rCnt++)
            {
                System.Array MyValues = (System.Array)xlWorkSheet.get_Range(Constants.START_COL + rCnt.ToString(), Constants.END_COL + rCnt.ToString()).Cells.Value;

                foreach (var a in MyValues)
                {
                    LOG.full(rCnt.ToString() + ": " + (a == null ? "--blank--" : a.ToString() + " - " + a.GetType()));
                }

                if (tc != null && (String)xlWorkSheet.get_Range(Constants.TEST_ACTION_COL + rCnt.ToString()).Value == "End")
                {
                    parsed++;
                    // We have full test case, now get cracking
                    if (testName != null && tc.getName() != testName)
                    {
                        LOG.normal("Skipping test '" + tc.getName() + "'");
                        skipped++;
                    }
                    else
                    {
                        LOG.normal("Running test '" + tc.getName() + "'");
                        run++;
                        if (tc.run())
                        {
                            LOG.normal("Test '" + tc.getName() + "' passed");
                            passed++;
                        }
                        else
                        {
                            LOG.normal("Test '" + tc.getName() + "' failed");
                            failed++;
                        }
                    }
                }
                else
                {
                    if (xlWorkSheet.get_Range(Constants.TEST_NAME_COL + rCnt.ToString()).Value != null)
                    {
                        tc = new TestCase(logLevel, (System.Array)xlWorkSheet.get_Range(Constants.TEST_NAME_COL + rCnt.ToString(), Constants.TEST_BASEURL_COL + rCnt.ToString()).Cells.Value);
                    }

                    // if we have action, etc
                    if (tc != null && (xlWorkSheet.get_Range(Constants.TEST_ACTION_COL + rCnt.ToString()).Value != null || xlWorkSheet.get_Range(Constants.TEST_ACTION_ON_ELEMENT_COL + rCnt.ToString()).Value != null))
                    {
                        tc.addOperation((System.Array)xlWorkSheet.get_Range(Constants.TEST_ACTION_COL + rCnt.ToString(), Constants.TEST_ARGUMENT_COL + rCnt.ToString()).Cells.Value);
                    }
                }
            }

            xlWorkBook.Close(true, null, null);
            xlApp.Quit();

            releaseObject(xlWorkSheet);
            releaseObject(xlWorkBook);
            releaseObject(xlApp);

            LOG.always("Tests Parsed: " + parsed + " Tests Run: " + run + " Tests Skipped: " + skipped);
            if (run > 0) LOG.always("Tests Passed: " + passed + " (" + passed * 100 / run + "%) Tests Failed: " + failed + " (" + failed * 100 / run + "%)");
        }

        static private bool parseArgs(string[] args)
        {
            for (int i = 0; i < args.Length; i++)
            {
                if (args[i] == "-f")
                {
                    if (args[i + 1] != null)
                    {
                        inputFile = args[i + 1];
                        LOG.always("Input File: '" + inputFile + "'");
                    }
                    else
                    {
                        usage();
                        return false;
                    }
                }
                if (args[i] == "-t")
                {
                    if (args[i + 1] != null)
                    {
                        testName = args[i + 1];
                        LOG.always("Test Name to run: '" + testName + "'");
                    }
                    else
                    {
                        usage();
                        return false;
                    }
                }
                if (args[i] == "-l")
                {
                    if (args[i + 1] != null)
                    {
                        logLevel = args[i + 1].ToLower();
                        if (logLevel == "verbose" || logLevel == "normal" || logLevel == "full" || logLevel == "none")
                        {
                            LOG.always("Log level: '" + logLevel + "'");
                            LOG.setLogLevel(logLevel);
                        }
                        else
                        {
                            usage();
                            return false;
                        }
                    }
                    else
                    {
                        usage();
                        return false;
                    }
                }
            }

            // validate we have enough to get going
            if (inputFile != null) return true;

            usage();

            return false;
        }

        private static void usage()
        {
            Console.WriteLine(Process.GetCurrentProcess().ProcessName);
            Console.WriteLine("\t -f <path to input excel file>");
            Console.WriteLine("\t [-t <test name>]");
            Console.WriteLine("\t [-l <none|normal|verbose|full>]");
        }

        static private void releaseObject(object obj)
        {
            try
            {
                System.Runtime.InteropServices.Marshal.ReleaseComObject(obj);
                obj = null;
            }
            catch (Exception ex)
            {
                obj = null;
                LOG.always("Unable to release the Object " + ex.ToString());
            }
            finally
            {
                GC.Collect();
            }
        }
    }

}