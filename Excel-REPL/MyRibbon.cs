using System;
using System.Runtime.InteropServices;
using System.Windows.Forms;
using System.Collections;
using ExcelDna.Integration.CustomUI;
using clojure.lang;
using ExcelDna.Integration;

[ComVisible(true)]
public class MyRibbon : ExcelRibbon
{
    public void OnButtonPressed(IRibbonControl control)
    {
        return;
        //ClojureExcel.MainClass.format_code.invoke();
        var app = NetOffice.ExcelApi.Application.GetActiveInstance();
        NetOffice.ExcelApi.Range sel = (NetOffice.ExcelApi.Range)app.Selection;

        System.Collections.Generic.List<String> inputArray = new System.Collections.Generic.List<String>();
        String[] input;

        var rawArray = sel.Value as Object[,];
        if (rawArray != null)
        {
            int a = rawArray.GetLowerBound(0);
            int b = rawArray.GetUpperBound(0);
            int c = rawArray.GetLowerBound(1);
            for (var i = a; i <= b; i++)
            {
                if (rawArray[i, c] != null)
                {
                    inputArray.Add(rawArray[i, c].ToString());
                }
                else
                {
                    //do we really want 1 million blank lines?
                    //inputArray.Add("");
                }
            }
            input = inputArray.ToArray();
        }
        else if (sel.Value != null)
        {
            input = new String[] { sel.Value.ToString() };
        }
        else
        {
            input = new String[] { "" };
        }

        //ClojureExcel.MainClass.format_code.invoke(input);
    }
    public static void ShowContents()
    {
        var app = NetOffice.ExcelApi.Application.GetActiveInstance();
        NetOffice.ExcelApi.Range sel = (NetOffice.ExcelApi.Range)app.Selection;
        foreach (var item in sel)
        {
            MessageBox.Show(item.Value.GetType().ToString());
        }

    }

    public static void SetOutput(IPersistentCollection output)
    {
        ISeq seq = output.seq();
        var app = NetOffice.ExcelApi.Application.GetActiveInstance();
        NetOffice.ExcelApi.Range sel = (NetOffice.ExcelApi.Range)app.Selection;

        var rawArray = sel.Value as Object[,];
        if (rawArray != null)
        {
            int a = rawArray.GetLowerBound(0);
            int b = rawArray.GetUpperBound(0);
            int c = rawArray.GetLowerBound(1);
            for (var i = a; i <= b; i++)
            {
                if (rawArray[i, c] != null)
                {
                    rawArray[i, c] = seq.first();
                    seq = seq.next();
                }
            }
            sel.Value = rawArray;
        }
        else
        {
            sel.Value = seq.first();
        }
    }

    public static void CopyHtmlToClipboard(String css, String html)
    {
        String template = ClojureExcel.MainClass.ResourceSlurp("SampleClipboard.txt");
        String pasted = String.Format(template, css, html);
        String startHtml = "StartHTML:0000000213";
        String endHtml = "EndHTML:0000002410";
        String startFragment = "StartFragment:0000000249";
        String endFragment = "EndFragment:0000002374";

        int startHtmli = pasted.IndexOf("<html>");
        int endHtmli = pasted.Length;
        int startFragmenti = pasted.IndexOf("<!--StartFragment-->");
        int endFragmenti = pasted.IndexOf("<!--EndFragment-->");

        String startHtml2 = String.Format("StartHTML:{0:D10}", startHtmli);
        String endHtml2 = String.Format("EndHTML:{0:D10}", endHtmli);
        String startFragment2 = String.Format("StartFragment:{0:D10}", startFragmenti);
        String endFragment2 = String.Format("EndFragment:{0:D10}", endFragmenti);

        String final = pasted.Replace(startHtml, startHtml2).Replace(endHtml, endHtml2).Replace(startFragment, startFragment2).Replace(endFragment, endFragment2);
        Clipboard.SetData(DataFormats.Html, final);
    }

    public static void PasteOutput()
    {
        var app = NetOffice.ExcelApi.Application.GetActiveInstance();
        NetOffice.ExcelApi.Range sel = (NetOffice.ExcelApi.Range)app.Selection;
        //ClojureExcel.MainClass.spit.invoke("test.txt", Clipboard.GetData(DataFormats.Html));
        //Clipboard.SetData(DataFormats.Html, ClojureExcel.MainClass.slurp.invoke("test.txt"));
        //Clipboard.SetData(DataFormats.Text, html);
        String html = clojureexcel.Python.Html("(def a 3)");
        String css = clojureexcel.Python.Css();
        CopyHtmlToClipboard(css, html);
        sel.PasteSpecial();
    }

    public static void ColorSelection(object[] data)
    {
        var app = NetOffice.ExcelApi.Application.GetActiveInstance();
        NetOffice.ExcelApi.Range sel = (NetOffice.ExcelApi.Range)app.Selection;

        for (var i = 0; i < data.Length; i += 3)
        {
            int j = (int)data[i];
            int k = (int)data[i + 1];
            double l = (double)data[i + 2];

            sel.Characters(j, k).Font.Color = l;
        }
    }

    private static String[] rawColors = {"#AAA", "#AAA", "#0F0", "#00FF7F", "#00FFFF", "#836FFF", "#FF00FF", "#9B30FF", "#00FF7F", "#00FFFF", "#836FFF", "#FF00FF", "#9B30FF"};
    private static double[] colors;

    private static void initColors()
    {
        colors = new double[rawColors.Length];
        for (var i = 0; i < rawColors.Length; i++)
        {
            colors[i] = ToDouble(System.Drawing.ColorTranslator.FromHtml(rawColors[i]));
        }
    }
    static MyRibbon()
    {
        //initColors();
    }

    public static void RandomColors(int j)
    {
        var app = NetOffice.ExcelApi.Application.GetActiveInstance();
        NetOffice.ExcelApi.Range sel = (NetOffice.ExcelApi.Range)app.Selection;
        Random r = new Random();

        for (var i = 0; i < j; i++)
        {
            System.Drawing.Color c = r.NextDouble() < 0.5 ? System.Drawing.Color.Red : System.Drawing.Color.Blue;
            sel.Characters(i, 1).Font.Color = colors[i];
        }
    }

    private static double ToDouble(System.Drawing.Color color)
    {
        uint returnValue = color.B;
        returnValue = returnValue << 8;
        returnValue += color.G;
        returnValue = returnValue << 8;
        returnValue += color.R;
        return returnValue;
    }
}