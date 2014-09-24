/*
 * Excelenium - An excel driven Selenium test suite
 * Copyright (c) Paul Connolly, paul.connolly@managed.info 2014
 * All rights reserved
 */

using System;

namespace Excelenium
{
    static class LOG
    {
        static private int logLevel;

        static public void setLogLevel(String level)
        {
            if(level == "none")
                logLevel = 0;
            else if(level == "normal")
                logLevel = 1;
            else if(level == "verbose")
                logLevel = 2;
            else if(level == "full")
                logLevel = 3;
        }
        
        static public void verbose(String s)
        {
            if (logLevel >= 2)
                log(s);
        }

        static public void normal(String s)
        {
            if (logLevel >= 1)
                log(s);
        }

        static public void full(String s)
        {
            if (logLevel >= 3)
                log(s);
        }

        static public void always(String s)
        {
            log(s);
        }

        static private void log(String s)
        {
            Console.Write(DateTime.Now + " - ");
            Console.WriteLine(s);
        }
    }
}