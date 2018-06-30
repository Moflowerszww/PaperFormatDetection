using System;
using System.Collections.Generic;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using PaperFormatDetection.Tools;

namespace PaperFormatDetection.Paperbase
{

    class Text
    {
        /// <summary>
        /// tTitle数组为正文中标题的属性
        /// tfulltext数组为正文中段落的属性
        /// 标题6个参数 顺序分别是 0缩进 1字体 2字号 3对齐方式 4段前间距 5段后间距 6行间距 7加粗
        /// 段落7个参数 顺序分别是 0缩进 1字体 2字号 3对齐方式 4段前间距 5段后间距 6行间距 7加粗
        /// </summary>
        protected string[,] tTitle = new string[3, 8];
        protected string[] tText = new string[8];
        static string[] NotesPattern = { "[-][-].*", "[/][*].*?", ".*?[*][/]", "[/][/].*", "[#＃].*", "<!--.*?-->", "[\"].+?[\"]", "[%].*", "[\'][\u4E00-\u9FA5]*[\']",
                                         "[^\u4E00-\u9FA5]*[-][-].*", "[^\u4E00-\u9FA5]*[/][/].*", "[^\u4E00-\u9FA5]*[%].*", "[^\u4E00-\u9FA5]*[#＃].*"};
        static string[] ProNumbering = { "[（].*?[）]", "[(].*?[)]", "[①②③④⑤⑥⑦⑧⑨⑩]", ".[）).．]", "[1-9一二三四五六七八九十][、]", "[◆●]" };
        bool ChineseNumbering = false;
        bool EnglishNumbering = false;
        bool isParagraphBreak = false;
        int count = 0;

        public Text()
        {

        }

        public void detectAllText(List<DocumentFormat.OpenXml.OpenXmlElement> list, WordprocessingDocument doc)
        {
            //记录当前标题文本
            string curTitle = null;
            //记录当前段落文本
            string paratext = null;
            //记录空段落
            List<string> content = getContent(doc);
            int listCount = list.Count;
            //foreach (DocumentFormat.OpenXml.OpenXmlElement p in list)
            string pre = null;
            for (int i = 0; i < listCount; i++)
            {
                if (list[i].GetType().ToString() != "DocumentFormat.OpenXml.Wordprocessing.Paragraph")
                    continue;
                Paragraph p = (Paragraph)list[i];
                if (paratext != null && paratext != "")
                    pre = paratext;
                if (itIsPic(p))
                    continue;
                if (i == listCount - 1)
                    continue;
                paratext = Util.getFullText(p).Trim();
                //有空行
                if (paratext.Length == 0)
                {
                    bool isFiltering = false;
                    if (i + 1 < listCount && list[i + 1].GetType().ToString() == "DocumentFormat.OpenXml.Wordprocessing.Paragraph")
                    {
                        Paragraph nextpar = (Paragraph)list[i + 1];
                        string nexttext = Util.getFullText(nextpar).Trim();
                        //下一行是表名
                        if (nexttext.IndexOf('，') == -1 && nexttext.Length > 0 && (Regex.IsMatch(nexttext, @"^[表][ ]*?[0-9]")))
                            isFiltering = true;
                        //下一行是图
                        if (itIsPic(nextpar))
                            isFiltering = true;
                        //下一行是章标题
                        if (Util.pageDic.ContainsKey(nexttext))
                            isFiltering = true;
                        //下一行也是空行
                        if (nexttext.Length == 0)
                            isFiltering = true;
                    }
                    if (i - 1 >= 0)
                    {
                        //上一行是表
                        if (list[i - 1].GetType().ToString() == "DocumentFormat.OpenXml.Wordprocessing.Table")
                            isFiltering = true;
                        else
                        {
                            Paragraph prepar = (Paragraph)list[i - 1];
                            string pretext = Util.getFullText(prepar).Trim();
                            //上一行是图名
                            if (pretext.IndexOf('，') == -1 && pretext.Length > 0 && (Regex.IsMatch(pretext, @"^[图][ ]*?[0-9]") || Regex.IsMatch(pretext, @"^Fig")))
                                isFiltering = true;
                        }
                    }
                    if (!isFiltering)
                        Util.printError("正文此段落后不应有空行   " + curTitle + "----" + (pre.Length > 12 ? pre.Substring(0, 12) : pre));
                }
                if (!Regex.IsMatch(paratext, @"\S"))
                {
                    continue;
                }
                if (isCode(paratext) || p.GetFirstChild<DocumentFormat.OpenXml.Math.OfficeMath>() != null || Regex.IsMatch(paratext, "[（(].+?[)）]$"))
                {

                }
                else if (isTitle(p, content, doc))
                {
                    curTitle = paratext;
                    detectTitle(p, doc);
                }
                else
                {
                    //过滤中文图名表名英文表名英文图名
                    if (paratext.IndexOf('，') == -1 && paratext.Length > 0 && (Regex.IsMatch(paratext, @"^[表][ ]*?[0-9]") || Regex.IsMatch(paratext, @"^Fig") || Regex.IsMatch(paratext, @"^Tab") || Regex.IsMatch(paratext, @"^[图][ ]*?[0-9]")))
                        continue;
                    if (Regex.IsMatch(paratext, @"[图表][ ]*[1-9][0-9]*[.][1-9][0-9]*[续\s]+") ||
                        Regex.IsMatch(paratext, @"[图表][ ]*[1-9][0-9]*[.][1-9][0-9]*") ||
                        Regex.IsMatch(paratext, @"Tab\.*\ *[1-9][0-9]*\.[1-9][0-9]*") ||
                        Regex.IsMatch(paratext, @"Fig\.*\ *[1-9][0-9]*\.[1-9][0-9]*"))
                        continue;
                    detectText(p, curTitle, doc);
                }
            }
        }

        public void detectTitle(Paragraph p, WordprocessingDocument doc)
        {
            Regex[] reg = new Regex[4];
            reg[0] = new Regex(@"[1-9]");
            reg[1] = new Regex(@"[1-9][0-9]*\.[1-9][0-9]*");
            reg[2] = new Regex(@"[1-9][0-9]*\.[1-9][0-9]*\..+?");
            reg[3] = new Regex(@"[1-9][0-9]*\.[1-9][0-9]*\.[1-9][0-9]*\.[1-9][0-9]*");
            string title = Tool.getFullText(p).Trim();

            IEnumerable<Run> run = p.Elements<Run>();
            int index;
            Match m = reg[3].Match(title);
            for (index = 3; index > -1; index--)
            {
                m = reg[index].Match(title);
                if (m.Success == true)
                    break;
            }
            if (index == 3)
            {
                Util.printError("正文章节标题序号错误，不应超过三级标题----" + title);
            }
            else if (index > -1)
            {

                //章标题另起一页
                if (index == 0)
                {
                    if (Util.pageDic.ContainsKey(title))
                    {
                        string page = Util.pageDic[title];
                        int index1 = page.IndexOf("_");
                        int index2 = page.LastIndexOf("_");
                        if (page.Substring(0, index1) == page.Substring(index1 + 1, index2 - index1 - 1))
                        {
                            Util.printError("正文章标题需另起一页（位于页首）" + "----" + title);
                        }
                    }
                }
                if (index == 0)
                {
                    if (Regex.Match(title, @"[1-9][0-9]*[.．、]").Success)
                        Util.printError("正文标题序号结尾不应加点号(.)----" + title);
                }
                if (title.IndexOf("   ") == m.Length)
                {
                    Util.printError("正文标题序号与内容之间应有两个空格----" + title);
                }
                else if (title.IndexOf("  ") != m.Length)
                {
                    Util.printError("正文标题序号与内容之间应有两个空格----" + title);
                }
                /* index = 0 对应一级标题
                    index = 1 对应二级标题
                    index = 2 对应三级标题 */
                if (!Util.correctIndentation(p, doc, tTitle[index, 0]))
                    Util.printError("正文标题缩进错误，应为总体缩进 " + tTitle[index, 0] + "字符：" + "----" + title);
                if (!Util.correctfonts(p, doc, tTitle[index, 1], "Cambria"))
                {
                    if (Util.paperType == "本科")
                        Util.printError("正文标题字体错误，应为" + tTitle[index, 1] + "----" + title);
                    else
                        Util.printError("正文标题字体错误，应为序号Cambria，中文" + tTitle[index, 1] + "----" + title);
                }
                if (!Util.correctsize(p, doc, tTitle[index, 2]))
                    Util.printError("正文标题字号错误，应为" + tTitle[index, 2] + "----" + title);
                if (!Util.correctJustification(p, doc, tTitle[index, 3]) && !Util.correctJustification(p, doc, "两端对齐"))
                    Util.printError("正文标题未" + tTitle[index, 3] + "----" + title);
                if (!Util.correctSpacingBetweenLines_Be(p, doc, tTitle[index, 4]))
                    Util.printError("正文标题段前距错误，应为" + Util.getLine(tTitle[index, 4]) + "----" + title);
                if (!Util.correctSpacingBetweenLines_Af(p, doc, tTitle[index, 5]))
                    Util.printError("正文标题段后距错误，应为" + Util.getLine(tTitle[index, 5]) + "----" + title);
                if (!Util.correctSpacingBetweenLines_line(p, doc, tTitle[index, 6]))
                    Util.printError("正文标题行间距错误，应为" + Convert.ToDouble(tTitle[index, 6]) / 240 + "倍行距" + "----" + title);
                if (!Util.correctBold(p, doc, Convert.ToBoolean(tTitle[index, 7])))
                    Util.printError("正文标题" + (Convert.ToBoolean(tTitle[index, 7]) ? "需要" : "不需") + "加粗" + "----" + title);
            }
            else
            {
                if (p.ParagraphProperties != null)
                {
                    if (p.ParagraphProperties.NumberingProperties == null)
                    {
                        Util.printError("正文标题序号缺失，应为阿拉伯数字----" + title);
                    }
                    else
                    {
                        index = p.ParagraphProperties.NumberingProperties.NumberingLevelReference.Val;
                        if (index < 3)
                        {
                            /* index = 0 对应一级标题
                            index = 1 对应二级标题
                            index = 2 对应三级标题 */
                            if (!Util.correctIndentation(p, doc, tTitle[index, 0]))
                                Util.printError("正文标题缩进错误，应为总体缩进 " + tTitle[index, 0] + "字符：" + "----" + title);
                            if (!Util.correctfonts(p, doc, tTitle[index, 1], "Cambria"))
                                Util.printError("正文标题字体错误，应为" + tTitle[index, 1] + "----" + title);
                            if (!Util.correctsize(p, doc, tTitle[index, 2]))
                                Util.printError("正文标题字号错误，应为" + tTitle[index, 2] + "----" + title);
                            if (!Util.correctJustification(p, doc, tTitle[index, 3]) && !Util.correctJustification(p, doc, "两端对齐"))
                                Util.printError("正文标题未" + tTitle[index, 3] + "----" + title);
                            if (!Util.correctSpacingBetweenLines_Be(p, doc, tTitle[index, 4]))
                                Util.printError("正文标题段前距错误，应为" + Util.getLine(tTitle[index, 4]) + "----" + title);
                            if (!Util.correctSpacingBetweenLines_Af(p, doc, tTitle[index, 5]))
                                Util.printError("正文标题段后距错误，应为" + Util.getLine(tTitle[index, 5]) + "----" + title);
                            if (!Util.correctSpacingBetweenLines_line(p, doc, tTitle[index, 6]))
                                Util.printError("正文标题行间距错误，应为" + Convert.ToDouble(tTitle[index, 6]) / 240 + "倍行距" + "----" + title);
                            if (!Util.correctBold(p, doc, Convert.ToBoolean(tText[7])))
                                Util.printError("正文标题" + (Convert.ToBoolean(tText[7]) ? "需要" : "不需") + "加粗" + "----" + title);
                        }
                        else
                        {
                            Util.printError("正文章节标题序号错误，不应超过三级标题" + "----" + title);
                        }
                    }
                }
            }
        }

        public void detectText(Paragraph p, string curTitle, WordprocessingDocument doc)
        {
            IEnumerable<Run> runList = p.Elements<Run>();
            string paratext = Util.getFullText(p).Trim();
            bool isPro = false;
            //特殊手动输入的项目符号检测
            if (Regex.IsMatch(p.InnerText, @"^\ eq \\o\\ac\(○,[0-9]\)"))
            {
                if (!paratext.StartsWith(" ") || paratext.StartsWith("  "))
                    Util.printError("正文项目编号与内容之间应有一个空格间隔：" + curTitle + "----" + (paratext.Length < 10 ? paratext : (paratext.Substring(0, 10) + "...")));
            }
            else if (Regex.IsMatch(p.InnerText, @"^\ eq \\o\\ac\(○,[^0-9]\)"))
            {
                Util.printError("正文项目编号错误，应为(1)(2)格式，子编号应为①②格式：" + curTitle + "----" + (paratext.Length < 10 ? paratext : (paratext.Substring(0, 10) + "...")));
            }
            //正常手动编号的项目符号检测
            isPro = detectProTitle(paratext, curTitle);
            //自动生成的项目符号检测
            if (p.ParagraphProperties != null)
            {
                if (p.ParagraphProperties.NumberingProperties != null)
                {
                    isPro = true;
                    string numberingId = p.ParagraphProperties.NumberingProperties.NumberingId.Val;
                    string ilvl = p.ParagraphProperties.NumberingProperties.NumberingLevelReference.Val;
                    Numbering numbering1 = doc.MainDocumentPart.NumberingDefinitionsPart.Numbering;
                    IEnumerable<NumberingInstance> nums = numbering1.Elements<NumberingInstance>();
                    IEnumerable<AbstractNum> abstractNums = numbering1.Elements<AbstractNum>();
                    foreach (NumberingInstance num in nums)
                    {
                        if (num.NumberID == numberingId)
                        {
                            Int32 abstractNumId1 = num.AbstractNumId.Val;
                            foreach (AbstractNum abstractNum in abstractNums)
                            {
                                if (abstractNum.AbstractNumberId == abstractNumId1)
                                {
                                    Level level = abstractNum.GetFirstChild<Level>();
                                    if (level.GetFirstChild<NumberingFormat>().Val == "decimalEnclosedCircle" || level.GetFirstChild<NumberingFormat>().Val == "decimalEnclosedCircleChinese")
                                    {
                                        //采用自动编号的二级项目编号
                                    }
                                    else if (level.GetFirstChild<NumberingFormat>().Val != "decimal")
                                    {
                                        Util.printError("正文项目编号错误，应为(1)(2)格式，子编号应为①②格式：" + curTitle + "----" + (paratext.Length < 10 ? paratext : (paratext.Substring(0, 10) + "...")));
                                    }
                                    else if (level.GetFirstChild<LevelText>().Val.InnerText != "（%1）" && level.GetFirstChild<LevelText>().Val.InnerText != "(%1)")
                                    {
                                        Util.printError("正文项目编号错误，应为(1)(2)格式，子编号应为①②格式：" + curTitle + "----" + (paratext.Length < 10 ? paratext : (paratext.Substring(0, 10) + "...")));
                                    }
                                    break;
                                }
                            }
                            break;
                        }
                    }
                }
            }
            if (runList != null && p.ParagraphProperties != null)
            {
                Match m = Regex.Match(Util.getFullText(p), @"\s+\S");
                if (!Util.correctIndentation(p, doc, tText[0]))
                {
                    if (Util.getFullText(p).IndexOf("    ") == 0)
                    {
                        m = Regex.Match(Util.getFullText(p).Substring(4), @"\s+\S");
                        if (m.Success && m.Index == 0 && !isPro)
                            Util.printError("正文段落前存在多余空格：" + curTitle + "----" + (paratext.Length < 10 ? paratext : (paratext.Substring(0, 10) + "...")));
                    }
                    else
                        Util.printError("正文段落缩进错误，应为左侧缩进0字符,首行缩进" + tText[0] + "字符：" + curTitle + "----" + (paratext.Length < 10 ? paratext : (paratext.Substring(0, 10) + "...")));
                }
                else
                {
                    if (m.Success && m.Index == 0 && !isPro)
                        Util.printError("正文段落前存在多余空格：" + curTitle + "----" + (paratext.Length < 10 ? paratext : (paratext.Substring(0, 10) + "...")));
                }
                if (!Util.correctfonts(p, doc, tText[1], "Times New Roman"))
                    Util.printError("正文段落字体错误，应为" + tText[1] + "：" + curTitle + "----" + (paratext.Length < 10 ? paratext : (paratext.Substring(0, 10) + "...")));
                if (!Util.correctsize(p, doc, tText[2]))
                {
                    if (isParagraphBreak)
                    {
                        isParagraphBreak = false;
                    }
                    else
                    {
                        if (isPro)
                            Util.printError("正文段落字号错误，应为" + tText[2] + "：" + curTitle + "----" + (paratext.Length < 10 ? paratext : (paratext.Substring(0, 10) + "...")));
                        else
                        {
                            m = Regex.Match(paratext.Substring(paratext.Length - 1, 1), @"[，、“；0-9a-zA-Z\u4E00-\u9FA5]");
                            if (m.Success)
                            {
                                isParagraphBreak = false;
                                Util.printError("正文段落字号错误，应为" + tText[2] + "：" + curTitle + "----" + (paratext.Length < 10 ? paratext : (paratext.Substring(0, 10) + "...")));
                            }
                            else isParagraphBreak = true;
                        }
                    }
                }
                //if (!Util.correctJustification(p, doc, tText[3]) && !Util.correctJustification(p, doc, "两端对齐"))
                //Util.printError("正文段落未" + tText[3] + "：" + curTitle + "----" + (paratext.Length < 10 ? paratext : (paratext.Substring(0, 10) + "...")));
                if (!Util.correctSpacingBetweenLines_Be(p, doc, tText[4]))
                    Util.printError("正文段落段前距应为" + Util.getLine(tText[4]) + "：" + curTitle + "----" + (paratext.Length < 10 ? paratext : (paratext.Substring(0, 10) + "...")));
                if (!Util.correctSpacingBetweenLines_Af(p, doc, tText[5]))
                    Util.printError("正文段落段后距应为" + Util.getLine(tText[5]) + "：" + curTitle + "----" + (paratext.Length < 10 ? paratext : (paratext.Substring(0, 10) + "...")));
                if (!Util.correctSpacingBetweenLines_line(p, doc, tText[6]))
                    Util.printError("正文段落行间距应为" + Convert.ToDouble(tText[6]) / 240 + "倍行距：" + curTitle + "----" + (paratext.Length < 10 ? paratext : (paratext.Substring(0, 10) + "...")));
                if (!Util.correctBold(p, doc, Convert.ToBoolean(tText[7])))
                    Util.printError("正文段落" + (Convert.ToBoolean(tText[7]) ? "需要" : "不需") + "加粗：" + curTitle + "----" + (paratext.Length < 10 ? paratext : (paratext.Substring(0, 10) + "...")));

            }
        }
        public static List<string> getContent(WordprocessingDocument doc)
        {
            Body body = doc.MainDocumentPart.Document.Body;
            IEnumerable<Paragraph> para = body.Elements<Paragraph>();
            List<string> list = new List<string>();
            bool begin = false;

            foreach (Paragraph p in para)
            {
                Run r = p.GetFirstChild<Run>();
                Hyperlink h = p.GetFirstChild<Hyperlink>();
                FieldChar f = null;
                String paratext = "";
                if (r != null)
                {
                    paratext = Util.getFullText(p).Trim();
                    f = r.GetFirstChild<FieldChar>();
                }
                else if (h != null)
                {
                    paratext = Util.getFullText(p);
                }
                else
                {
                    continue;
                }
                if (paratext.Replace(" ", "") == "目录")
                {
                    begin = true;
                    continue;
                }
                if (begin)
                {
                    if (f == null && h == null)
                    {
                        return list;
                    }
                    else
                    {
                        list.Add(paratext);
                    }
                }
            }
            return list;
        }
        //判断是否为代码，用于正文中的代码过滤
        public static bool isCode(string str)
        {
            Match m = Regex.Match(str, ProNumbering[0]);
            for (int i = 0; i < ProNumbering.Length; i++)
            {
                m = Regex.Match(str, ProNumbering[i]);
                if (m.Success && m.Index == 0) return false;
            }

            m = Regex.Match(str, @"[\u4E00-\u9FA5]");
            if (!m.Success)
                return true;
            else
            {
                m = Regex.Match(str, @NotesPattern[0]);
                if (m.Success && m.Index == 0)
                    return true;
                for (int i = 1; i < NotesPattern.Length; i++)
                {
                    m = Regex.Match(str, @NotesPattern[i]);
                    if (m.Success)
                        return true;
                }
            }
            return false;
        }

        public bool isTitle(Paragraph p, List<string> content, WordprocessingDocument doc)
        {
            bool b = false;
            int index, counter = 0;
            string title = Util.getFullText(p).Trim();
            Regex[] reg = new Regex[4];
            //一级标题
            reg[0] = new Regex(@"[1-9][0-9]*[\s]+?");
            reg[1] = new Regex(@"[1-9][0-9]*[.．、\u4E00-\u9FA5]");
            //二级标题
            reg[2] = new Regex(@"[1-9][0-9]*\.[1-9][0-9]*");
            //三级标题
            reg[3] = new Regex(@"[1-9][0-9]*\.[1-9][0-9]*\.[1-9][0-9]*");


            if (countWords(title) < 50)
            {
                Match m = reg[3].Match(title);

                for (index = 3; index > -1; index--)
                {
                    m = reg[index].Match(title);
                    if (m.Success == true)
                        break;
                }
                if (m.Index == 0)
                {
                    if (index > -1)
                    {
                        if (index == 1)
                        {
                            foreach (string s in content)
                            {
                                if (s.Contains(title) && Util.correctsize(p, doc, tTitle[0, 1]))
                                {
                                    b = true;
                                    break;
                                }
                            }
                        }
                        else
                            b = true;
                    }
                    else
                    {
                        if (p.ParagraphProperties != null)
                        {
                            if (p.ParagraphProperties.NumberingProperties == null)
                            {
                                foreach (string s in content)
                                {
                                    if (s.Contains(title))
                                    {
                                        double similarity = title.Length / s.Length;
                                        if (similarity > 0.8)
                                        {
                                            b = true;
                                            break;
                                        }
                                    }
                                }
                            }
                            else
                            {
                                string numberingId = p.ParagraphProperties.NumberingProperties.NumberingId.Val;
                                Numbering numbering = doc.MainDocumentPart.NumberingDefinitionsPart.Numbering;
                                IEnumerable<NumberingInstance> nums = numbering.Elements<NumberingInstance>();
                                IEnumerable<AbstractNum> abstractNums = numbering.Elements<AbstractNum>();
                                foreach (NumberingInstance num in nums)
                                {
                                    if (num.NumberID == numberingId)
                                    {
                                        Int32 abstractNumId1 = num.AbstractNumId.Val;
                                        foreach (AbstractNum abstractNum in abstractNums)
                                        {
                                            if (abstractNum.AbstractNumberId == abstractNumId1)
                                            {
                                                IEnumerable<Level> levels = abstractNum.Elements<Level>();
                                                foreach (Level level in levels)
                                                {
                                                    counter = counter + 1;
                                                    if (counter == 1)
                                                    {
                                                        if (level.GetFirstChild<LevelText>().Val.InnerText != "%1." && level.GetFirstChild<LevelText>().Val.InnerText != "%1")
                                                            break;
                                                    }
                                                    else if (counter == 2)
                                                    {
                                                        if (level.GetFirstChild<LevelText>().Val.InnerText != "%1.%2")
                                                        {
                                                            foreach (string s in content)
                                                            {
                                                                double similarity = title.Length / s.Length;
                                                                if (similarity > 0.8)
                                                                {
                                                                    b = true;
                                                                    break;
                                                                }
                                                            }
                                                            break;
                                                        }
                                                    }
                                                    else if (counter == 3)
                                                    {
                                                        if (level.GetFirstChild<LevelText>().Val.InnerText != "%1.%2.%3")
                                                            break;
                                                        b = true;
                                                    }
                                                    else
                                                        break;
                                                }
                                                break;
                                            }
                                        }
                                        break;
                                    }
                                }
                            }
                        }
                    }
                }
            }
            return b;
        }

        public bool detectProTitle(string paratext, string curTitle)
        {
            bool isPro = false;
            Match m = Regex.Match(paratext, "^[（].+?[）]");
            if (m.Success && m.Length < 5)
            {
                isPro = true;
                ChineseNumbering = true;
                if (count < 1 && EnglishNumbering)
                {
                    Util.printError("正文项目编号格式错误，应为中文圆括号：" + curTitle + "----" + (paratext.Length < 10 ? paratext : (paratext.Substring(0, 10) + "...")));
                    count++;
                }
                if (!Regex.Match(m.Value.Substring(1, m.Length - 2), "[0-9][1-9]*").Success)
                    Util.printError("正文项目编号错误，应为(1)(2)格式，子编号应为①②格式：" + curTitle + "----" + (paratext.Length < 10 ? paratext : (paratext.Substring(0, 10) + "...")));
                if (paratext.Substring(m.Length, 1) != " " || paratext.Substring(m.Length, 2) == "  ")
                    Util.printError("正文项目编号与内容之间应有一个空格间隔：" + curTitle + "----" + (paratext.Length < 10 ? paratext : (paratext.Substring(0, 10) + "...")));
            }
            m = Regex.Match(paratext, "^[(].+?[)]");
            if (m.Success && m.Length < 5)
            {
                isPro = true;
                EnglishNumbering = true;
                if (count < 1 && ChineseNumbering)
                {
                    Util.printError("正文项目编号格式错误，应为中文圆括号：" + curTitle + "----" + (paratext.Length < 10 ? paratext : (paratext.Substring(0, 10) + "...")));
                    count++;
                }
                if (!Regex.Match(m.Value.Substring(1, m.Length - 2), "[0-9][1-9]*").Success)
                    Util.printError("正文项目编号错误，应为(1)(2)格式，子编号应为①②格式：" + curTitle + "----" + (paratext.Length < 10 ? paratext : (paratext.Substring(0, 10) + "...")));
                if (paratext.Substring(m.Length, 1) != " " || paratext.Substring(m.Length, 2) == "  ")
                    Util.printError("正文项目编号与内容之间应有一个空格间隔：" + curTitle + "----" + (paratext.Length < 10 ? paratext : (paratext.Substring(0, 10) + "...")));
            }
            m = Regex.Match(paratext, "^[①②③④⑤⑥⑦⑧⑨⑩]");
            if (m.Success)
            {
                isPro = true;
                if (paratext.Substring(1, 1) != " " || paratext.Substring(1, 2) == "  ")
                    Util.printError("正文项目编号与内容之间应有一个空格间隔：" + curTitle + "----" + (paratext.Length < 10 ? paratext : (paratext.Substring(0, 10) + "...")));
            }
            m = Regex.Match(paratext, "^.[）).．]");
            if (m.Success)
            {
                isPro = true;
                if (!Regex.IsMatch(paratext, @"[A-Z][.]\s*[A-Z][a-zA-Z]"))
                {
                    Util.printError("正文项目编号错误，应为(1)(2)格式，子编号应为①②格式：" + curTitle + "----" + (paratext.Length < 10 ? paratext : (paratext.Substring(0, 10) + "...")));
                }
            }
            m = Regex.Match(paratext, "^[1-9一二三四五六七八九十][、丶]");
            if (m.Success)
            {
                isPro = true;
                Util.printError("正文项目编号错误，应为(1)(2)格式，子编号应为①②格式：" + curTitle + "----" + (paratext.Length < 10 ? paratext : (paratext.Substring(0, 10) + "...")));
            }
            m = Regex.Match(paratext, "^[◆●]");
            if (m.Success)
            {
                isPro = true;
                Util.printError("正文项目编号错误，应为(1)(2)格式，子编号应为①②格式：" + curTitle + "----" + (paratext.Length < 10 ? paratext : (paratext.Substring(0, 10) + "...")));
            }
            return isPro;
        }

        public int countWords(string s)
        {
            if (Regex.IsMatch(s, @"[\u4E00-\u9FA5]"))
            {
                return s.Length;
            }
            else
            {
                string[] strArray = s.Split(' ');
                return strArray.Length;
            }
        }

        public bool itIsPic(Paragraph p)
        {
            IEnumerable<Run> runlist = p.Elements<Run>();
            foreach (Run r in runlist)
            {
                if (r != null)
                {
                    Drawing d = r.GetFirstChild<Drawing>();
                    Picture pic = r.GetFirstChild<Picture>();
                    EmbeddedObject objects = r.GetFirstChild<EmbeddedObject>();
                    AlternateContent Alt = r.GetFirstChild<AlternateContent>();
                    if (d != null || pic != null || objects != null)
                    {
                        return true;
                    }
                    else if (Alt != null)
                    {
                        AlternateContentChoice AltChoice = Alt.GetFirstChild<AlternateContentChoice>();
                        AlternateContentFallback AltFallback = Alt.GetFirstChild<AlternateContentFallback>();
                        if (AltChoice != null && AltFallback != null)
                        {
                            if ((AltChoice.GetFirstChild<Drawing>() != null || AltFallback.GetFirstChild<Picture>() != null))
                            {
                                return true;
                            }
                        }
                    }
                }
            }
            return false;
        }
    }
}
