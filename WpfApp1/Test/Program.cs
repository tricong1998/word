using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Word;
namespace Test
{
    class Program
    {
        //static int  valueFalse = 9999999;

        static void Main(string[] args)
        {
            Words w = new Words();
            Application application = new Application();
            application.Visible = true;
            Document document;
            //ValidateAnswer
            string path = @"E:\file.docx";
            string path1 = @"E:\path2 - Copy.docx";
            string path2 = @"E:\path2.docx";
            
            if (w.compare(path1, path2))
            {
                Console.WriteLine("giong");
            }
            else
            {
                Console.WriteLine("khac");
            }
            Console.ReadLine();

            Console.ReadKey();
            //document = application.Documents.Open(path.Trim());
            //Range range = document.Range();
            //int a = range.Start;
            //int b = range.End;
            ////var x = range.Text.Trim().Equals("");
            ////var y = range.Font.Name.ToString().Trim();
            //List<string> texts = new List<string>();
            //List<Range> ranges = w.classifyRange(document);
            //for (int i = 0; i < ranges.Count; i++)
            //{
            //    string text = ranges[i].Text;
            //    texts.Add(text);
            //}
            //Range range2 = document.Range();
            //range2.Start = 23;
            //range2.End = 24;
            //Range rangeTemp = ranges[0];
            //rangeTemp.End = rangeTemp.Start + 1;
            //Range rangeTemp2 = document.Range() ;
            //rangeTemp2.Start = 11;
            //rangeTemp2.End = 12;
            //var b1 = rangeTemp.Font.Line.ForeColor.Brightness;
            //var b2 = rangeTemp2.Font.Line.ForeColor.Brightness;
            //string textRange = rangeTemp.Text;
            //string textRange2 = rangeTemp2.Text;
            //string text0 = texts[0];
            //string text1 = texts[1];
            //string text2 = texts[2];
            //string text3 = texts[3];
            //string text4 = texts[4];
            //string text5 = texts[5];
            //string text6 = texts[6];
            //string text7 = texts[7];
            //string text8 = texts[8];
            ////        float po1 = rangeTemp.Font.Fill.ForeColor.Brightness;
            ////        float po2 = rangeTemp2.Font.Fill.ForeColor.Brightness;
            ////        float po1_1 = rangeTemp.Font.Line.ForeColor.Brightness;
            ////        float po2_1 = rangeTemp.Font.Line.ForeColor.Brightness;
            ////        int p0 = rangeTemp.Font.Fill.ForeColor.RGB;
            ////        string s0 = rangeTemp.Font.Fill.ForeColor.ObjectThemeColor.ToString();
            ////        int p1 = rangeTemp2.Font.Fill.ForeColor.RGB;
            ////        string s1 = rangeTemp2.Font.Fill.ForeColor.ObjectThemeColor.ToString();
            ////        int p0_0 = rangeTemp.Font.Fill.BackColor.RGB;
            //////        string s0_0 = rangeTemp.Font.Fill.BackColor.ObjectThemeColor.ToString();
            ////        int p1_1 = rangeTemp2.Font.Fill.BackColor.RGB;
            //// //       string s1_1 = rangeTemp2.Font.Fill.BackColor.ObjectThemeColor.ToString();
            ////        int p0_1 = rangeTemp.Font.Line.BackColor.RGB;
            //// //       string s0_1 = rangeTemp.Font.Line.BackColor.ObjectThemeColor.ToString();
            ////        int p1_2 = rangeTemp2.Font.Line.BackColor.RGB;
            ////  //      string s1_2 = rangeTemp2.Font.Line.BackColor.ObjectThemeColor.ToString();
            ////        int p2 = rangeTemp.Font.Line.ForeColor.RGB;
            ////        string s2 = rangeTemp.Font.Line.ForeColor.ObjectThemeColor.ToString();
            ////        int p3 = rangeTemp2.Font.Line.ForeColor.RGB;
            ////        string s3 = rangeTemp2.Font.Line.ForeColor.ObjectThemeColor.ToString();
            ////        var p23 = ranges[1].Font.TextColor.ObjectThemeColor.ToString();
            //document.Close();
            //application.Quit();
            //Console.WriteLine(a);

        }
        private static String HexConverter(System.Drawing.Color c)
        {
            return "#" + c.R.ToString("X2") + c.G.ToString("X2") + c.B.ToString("X2");
        }

        private static String RGBConverter(System.Drawing.Color c)
        {
            return "RGB(" + c.R.ToString() + "," + c.G.ToString() + "," + c.B.ToString() + ")";
        }
    }
}
