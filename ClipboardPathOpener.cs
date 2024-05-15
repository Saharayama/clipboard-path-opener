using System;
using System.Diagnostics;
using System.IO;
using System.Text.RegularExpressions;
using System.Windows.Forms;

public class ClipboardPathOpener {
    [STAThread]
    public static void Main()
    {
        string clipboardText = Clipboard.GetText();
        foreach(var line in clipboardText.Split(new[] { "\r\n", "\r", "\n" }, StringSplitOptions.None))
        {
            if(!string.IsNullOrWhiteSpace(line))
            {
                OpenPath(FormatPath(line));
            }
        }
    }

    static string FormatPath(string p)
    {
        string path = "";
        p = Regex.Replace(p, "\"|^ *-|^\t*-|<|>", "").Trim();

        if(p.Contains(":") && p.Contains("/"))
        {
            path = p.Replace("/", "\\");
        }
        else if(p.Contains("/"))
        {
            path = "\\\\" + p.Replace("/", "\\");
        }
        else if(p.Contains("\\\\"))
        {
            p = Regex.Replace(p, "^\\\\+", "");
            path = "\\\\" + p;
        }
        else
        {
            path = p;
        }

        return path;
    }

    static void OpenPath(string pathFormatted)
    {
        string pathSuffix = Path.GetExtension(pathFormatted);
        if(pathSuffix.Contains("--"))
        {
            return;
        }

        string officePath = "C:\\Program Files\\Microsoft Office\\root\\Office16";

        if(File.Exists(pathFormatted) || Directory.Exists(pathFormatted))
        {
            if(pathSuffix.Contains("xl"))
            {
                var excelCmd = new ProcessStartInfo();
                excelCmd.FileName = "\"" + officePath + "\\EXCEL.EXE\"";
                excelCmd.Arguments = "/x " + pathFormatted;
                try
                {
                    Process.Start(excelCmd);
                }
                catch(System.ComponentModel.Win32Exception)
                {
                    return;
                }
            }
            else if(pathSuffix.Contains("doc"))
            {
                var wordCmd = new ProcessStartInfo();
                wordCmd.FileName = "\"" + officePath + "\\WINWORD.EXE\"";
                wordCmd.Arguments = "/w " + pathFormatted;
                try
                {
                    Process.Start(wordCmd);
                }
                catch(System.ComponentModel.Win32Exception)
                {
                    return;
                }
            }
            else
            {
                try
                {
                    Process.Start(new ProcessStartInfo() { FileName = pathFormatted });
                }
                catch(System.ComponentModel.Win32Exception)
                {
                    return;
                }
            }
        }
    }
}
