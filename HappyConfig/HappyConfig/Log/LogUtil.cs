using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

public class LogUtil
{
    public static void LogDebug(string message, params string [] args)
    {
        string val = string.Format(message, args);
        Console.WriteLine(val);
    }
}
