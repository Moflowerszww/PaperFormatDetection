using System;
using System.IO;
using PaperFormatDetection.Tools;
using System.Text.RegularExpressions;

namespace PaperFormatDetection.Frame
{
    public class Program
    {
        public static int Main(string[] args)
        {
            Util.paperType = "本科";
            //Util.paperType = "硕士";
            //Util.paperType = "博士";
            DateTime start = DateTime.Now;
            Util.paperPath = Util.environmentDir + "\\Papers\\基于哈希编码的网络流量分类方法的研究.docx";
            
            if (args.Length > 0)
                Util.paperPath = args[0];
            //0学术型硕士 1专业学位硕士 2本科 3博士
            if (args.Length > 1)
            {
                if (args[1] == "0")
                {
                    Util.paperType = "硕士";
                    Util.masterType = "学术型硕士";
                }
                else if (args[1] == "1")
                {
                    Util.paperType = "硕士";
                    Util.masterType = "专业学位硕士";
                }
                else if (args[1] == "2")
                {
                    Util.paperType = "本科";
                }
                else
                {
                    Util.paperType = "博士";
                }
            }
            if (args.Length > 2)
                Util.environmentDir = args[2];
            //获取页码
            Console.WriteLine("正在获取页码...");
            MSWord msword = new MSWord();
            if (Util.paperPath.EndsWith(".doc"))
            {
                Console.WriteLine("正在将doc文件转为docx...");
                Util.paperPath = msword.DocToDocx(Util.paperPath);
                Console.WriteLine("文件转换成功！");
            }
            Util.pageDic = msword.getPage(Util.paperPath);
            foreach (var item in Util.pageDic)
            {
                //Util.printError(item.Key + "  " + item.Value);
                Console.WriteLine(item.Key + "  " + item.Value);
            }
            Console.WriteLine("成功获取页码信息！");

            Undergraduate.PaperDetection UndergraduatePD = null;
            Master.PaperDetection MasterPD = null;
            Doctor.PaperDetection DoctorPD = null;
            if (Util.paperType.Equals("本科"))
                UndergraduatePD = new Undergraduate.PaperDetection(Util.paperPath);
            else if (Util.paperType.Equals("硕士"))
                MasterPD = new Master.PaperDetection(Util.paperPath);
            else if (Util.paperType.Equals("博士"))
                DoctorPD = new Doctor.PaperDetection(Util.paperPath);

            DateTime end = DateTime.Now;
            TimeSpan ts = end - start;
            Console.WriteLine("");
            Console.WriteLine(" <= 检测用时： " + ts.TotalSeconds + " =>");
            //Console.ReadKey();
            return 0;
        }
    }
}