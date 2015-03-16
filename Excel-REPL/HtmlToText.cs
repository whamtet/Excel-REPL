using HtmlAgilityPack;
using System;
using System.IO;
using System.Text.RegularExpressions;

public static class HtmlToText
{

    public static string Convert(string path)
    {
        HtmlDocument doc = new HtmlDocument();
        doc.Load(path);
        return ConvertDoc(doc);
    }

    public static string ConvertHtml(string html)
    {
        HtmlDocument doc = new HtmlDocument();
        doc.LoadHtml(html);
        return ConvertDoc(doc);
    }

    public static string ConvertDoc(HtmlDocument doc)
    {
        using (StringWriter sw = new StringWriter())
        {
            ConvertTo(doc.DocumentNode, sw);
            sw.Flush();
            return sw.ToString();
        }
    }

    internal static void ConvertContentTo(HtmlNode node, TextWriter outText, PreceedingDomTextInfo textInfo)
    {
        foreach (HtmlNode subnode in node.ChildNodes)
        {
            ConvertTo(subnode, outText, textInfo);
        }
    }
    public static void ConvertTo(HtmlNode node, TextWriter outText)
    {
        ConvertTo(node, outText, new PreceedingDomTextInfo(false));
    }
    internal static void ConvertTo(HtmlNode node, TextWriter outText, PreceedingDomTextInfo textInfo)
    {
        string html;
        switch (node.NodeType)
        {
            case HtmlNodeType.Comment:
                // don't output comments
                break;
            case HtmlNodeType.Document:
                ConvertContentTo(node, outText, textInfo);
                break;
            case HtmlNodeType.Text:
                // script and style must not be output
                string parentName = node.ParentNode.Name;
                if ((parentName == "script") || (parentName == "style"))
                {
                    break;
                }
                // get text
                html = ((HtmlTextNode)node).Text;
                // is it in fact a special closing node output as text?
                if (HtmlNode.IsOverlappedClosingElement(html))
                {
                    break;
                }
                // check the text is meaningful and not a bunch of whitespaces
                if (html.Length == 0)
                {
                    break;
                }
                if (!textInfo.WritePrecedingWhiteSpace || textInfo.LastCharWasSpace)
                {
                    html = html.TrimStart();
                    if (html.Length == 0) { break; }
                    textInfo.IsFirstTextOfDocWritten.Value = textInfo.WritePrecedingWhiteSpace = true;
                }
                outText.Write(HtmlEntity.DeEntitize(Regex.Replace(html.TrimEnd(), @"\s{2,}", " ")));
                if (textInfo.LastCharWasSpace = char.IsWhiteSpace(html[html.Length - 1]))
                {
                    outText.Write(' ');
                }
                break;
            case HtmlNodeType.Element:
                string endElementString = null;
                bool isInline;
                bool skip = false;
                int listIndex = 0;
                switch (node.Name)
                {
                    case "nav":
                        skip = true;
                        isInline = false;
                        break;
                    case "body":
                    case "section":
                    case "article":
                    case "aside":
                    case "h1":
                    case "h2":
                    case "header":
                    case "footer":
                    case "address":
                    case "main":
                    case "div":
                    case "p": // stylistic - adjust as you tend to use
                        if (textInfo.IsFirstTextOfDocWritten)
                        {
                            outText.Write("\r\n");
                        }
                        endElementString = "\r\n";
                        isInline = false;
                        break;
                    case "br":
                        outText.Write("\r\n");
                        skip = true;
                        textInfo.WritePrecedingWhiteSpace = false;
                        isInline = true;
                        break;
                    case "a":
                        if (node.Attributes.Contains("href"))
                        {
                            string href = node.Attributes["href"].Value.Trim();
                            if (node.InnerText.IndexOf(href, StringComparison.InvariantCultureIgnoreCase) == -1)
                            {
                                endElementString = "<" + href + ">";
                            }
                        }
                        isInline = true;
                        break;
                    case "li":
                        if (textInfo.ListIndex > 0)
                        {
                            outText.Write("\r\n{0}.\t", textInfo.ListIndex++);
                        }
                        else
                        {
                            outText.Write("\r\n*\t"); //using '*' as bullet char, with tab after, but whatever you want eg "\t->", if utf-8 0x2022
                        }
                        isInline = false;
                        break;
                    case "ol":
                        listIndex = 1;
                        goto case "ul";
                    case "ul": //not handling nested lists any differently at this stage - that is getting close to rendering problems
                        endElementString = "\r\n";
                        isInline = false;
                        break;
                    case "img": //inline-block in reality
                        if (node.Attributes.Contains("alt"))
                        {
                            outText.Write('[' + node.Attributes["alt"].Value);
                            endElementString = "]";
                        }
                        if (node.Attributes.Contains("src"))
                        {
                            outText.Write('<' + node.Attributes["src"].Value + '>');
                        }
                        isInline = true;
                        break;
                    default:
                        isInline = true;
                        break;
                }
                if (!skip && node.HasChildNodes)
                {
                    ConvertContentTo(node, outText, isInline ? textInfo : new PreceedingDomTextInfo(textInfo.IsFirstTextOfDocWritten) { ListIndex = listIndex });
                }
                if (endElementString != null)
                {
                    outText.Write(endElementString);
                }
                break;
        }
    }
}
internal class PreceedingDomTextInfo
{
    public PreceedingDomTextInfo(BoolWrapper isFirstTextOfDocWritten)
    {
        IsFirstTextOfDocWritten = isFirstTextOfDocWritten;
    }
    public bool WritePrecedingWhiteSpace { get; set; }
    public bool LastCharWasSpace { get; set; }
    public readonly BoolWrapper IsFirstTextOfDocWritten;
    public int ListIndex { get; set; }
}
internal class BoolWrapper
{
    public BoolWrapper() { }
    public bool Value { get; set; }
    public static implicit operator bool(BoolWrapper boolWrapper)
    {
        return boolWrapper.Value;
    }
    public static implicit operator BoolWrapper(bool boolWrapper)
    {
        return new BoolWrapper { Value = boolWrapper };
    }
}