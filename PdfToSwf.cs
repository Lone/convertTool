using System;
using System.Diagnostics;


/// <summary>
/// PdfToSwf 的摘要说明
/// </summary>
public class PdfToSwf
{
	public PdfToSwf()
	{
		//
		// TODO: 在此处添加构造函数逻辑
		//
	}

    public bool Convert(string filename)
    {
        if (filename == "")
        {
            return false;
        }

        string exePath = Environment.CurrentDirectory;

        string swfname = filename.Replace(".pdf", "_view");
        
        Process swftool = new Process();

        swftool.StartInfo.FileName = @exePath + "\\pdf2swf.exe";
        swftool.StartInfo.Arguments = " -t \"" + @filename + "\" -o \"" + @swfname + "\" -s flashversion=9 ";
        swftool.StartInfo.CreateNoWindow = true;
        swftool.StartInfo.UseShellExecute = true;
        swftool.StartInfo.WindowStyle = ProcessWindowStyle.Hidden;

        bool isStart = swftool.Start();
        swftool.WaitForExit();
        swftool.Close();

        return true;
    }
}
