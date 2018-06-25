using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Word;
namespace WpfApp1
{
    class Words
    {
        //so sanh 2 file document tu 2 duong dan truyen vao
        public bool compare(String filePathAnswer, String filePathCorrectAnswer)
        {
            Application app = new Application();
            Document answer;
            try
            {
                answer = app.Documents.Open(filePathAnswer.Trim());
            }
            catch (Exception e)
            {
                return false;
            }
            Document correctAnswer = app.Documents.Open(filePathCorrectAnswer.Trim());
            bool isCorrectAnswer = false;
            if (answer.Range().Text.Equals(correctAnswer.Range().Text))
            {
                List<Range> customRangesCorrect = classifyRange(correctAnswer);
                List<Range> customRangesAnswer = classifyRange(answer);
                if (customRangesCorrect.Count() == customRangesAnswer.Count())
                {
                    for (int i = 0; i < customRangesCorrect.Count(); i++)
                    {
                        if (!checkEqualRange(customRangesAnswer[i], customRangesCorrect[i]))
                        {
                            break;
                        }
                        if (i == customRangesCorrect.Count() - 1)
                        {
                            isCorrectAnswer = true;
                        }
                    }
                }
            }
            answer.Close();
            correctAnswer.Close();
            app.Quit();
            return isCorrectAnswer;
        }
        //tach document thanh cac range co dinh dang khac nhau, tra ve list cac range
        public List<Range> classifyRange(Document document)
        {
            List<Range> ranges = new List<Range>();
            Range range = document.Range();
            Range rangeTemp = document.Range();
            rangeTemp.Start = range.Start;
            rangeTemp.End = range.Start + 1;
            ranges.Add(rangeTemp);
            while (ranges[ranges.Count - 1].End < range.End)
            {
                Range customRange = ranges[ranges.Count - 1];
                Range newRange = document.Range();
                newRange.Start = customRange.End;
                newRange.End = customRange.End + 1;
                if (checkEqualRange(newRange, customRange))
                {
                    ranges[ranges.Count - 1].End++;
                }
                else
                {
                    ranges.Add(newRange);
                }
            }

            return ranges;
        }
        //so sanh 2 range
        public bool checkEqualRange(Range range1, Range range2)
        {
            if (checkEqualFont(range1, range2)
                && checkEqualParagraph(range1.ParagraphFormat, range2.ParagraphFormat)
                && checkEqualBorder(range1.Borders, range2.Borders)
                && checkEqualPageSetup(range1.PageSetup, range2.PageSetup)
                //&& checkTextEffect(range1, range2)
                )
            {
                return true;
            }
            else
            {
                return false;
            }
        }
        //so sanh font cua 2 range
        public bool checkEqualFont(Range range1, Range range2)
        {
            Font font1 = range1.Font;
            Font font2 = range2.Font;
            if (font1.Bold == font2.Bold
                && font1.Italic == font2.Italic
                && font1.Size == font2.Size
                && font1.Name == font2.Name
                && font1.Color == font2.Color
                && font1.StrikeThrough == font2.StrikeThrough
                && font1.UnderlineColor == font2.UnderlineColor
                && range1.Underline == range2.Underline
                && range1.HighlightColorIndex == range2.HighlightColorIndex
               )
            {
                return true;
            }
            else
            {
                return false;
            }
        }
        //so sanh dinh dang doan van ban
        public bool checkEqualParagraph(ParagraphFormat paragraphFormat1,
            ParagraphFormat paragraphFormat2)
        {
            if (paragraphFormat1.Alignment == paragraphFormat2.Alignment
                && paragraphFormat1.LeftIndent == paragraphFormat2.LeftIndent
                && paragraphFormat1.RightIndent == paragraphFormat2.RightIndent
                && paragraphFormat1.FirstLineIndent == paragraphFormat2.FirstLineIndent
                && paragraphFormat1.MirrorIndents == paragraphFormat2.MirrorIndents
                && paragraphFormat1.LineSpacingRule == paragraphFormat2.LineSpacingRule
                && paragraphFormat1.SpaceAfter == paragraphFormat2.SpaceAfter
                && paragraphFormat1.SpaceBefore == paragraphFormat2.SpaceBefore
                && paragraphFormat1.LineSpacing == paragraphFormat2.LineSpacing
                )
            {
                return true;
            }
            else
            {
                return false;
            }
        }
        //so sanh le trang
        public bool checkEqualBorder(Borders border1, Borders border2)
        {
            if (border1.OutsideLineStyle == border2.OutsideLineStyle
                && border1.OutsideColorIndex == border2.OutsideColorIndex
                && border1.OutsideLineWidth == border2.OutsideLineWidth
                && border1.DistanceFromBottom == border2.DistanceFromBottom
                && border1.DistanceFromLeft == border2.DistanceFromLeft
                && border1.DistanceFromRight == border2.DistanceFromRight
                && border1.DistanceFromTop == border1.DistanceFromTop)
            {
                return true;
            }
            else
            {
                return false;
            }
        }
        //so sanh can chinh trang
        public bool checkEqualPageSetup(PageSetup pageSetup1, PageSetup pageSetup2)
        {
            if (pageSetup1.LeftMargin == pageSetup2.LeftMargin
                && pageSetup1.RightMargin == pageSetup2.RightMargin
                && pageSetup1.BottomMargin == pageSetup2.BottomMargin
                && pageSetup1.TopMargin == pageSetup2.TopMargin
                && pageSetup1.PageHeight == pageSetup2.PageHeight
                && pageSetup1.PageWidth == pageSetup2.PageWidth)
            {
                return true;
            }
            else
            {
                return false;
            }
        }
        //so sanh hieu ung chu
        public bool checkTextEffect(Range range1, Range range2)
        {
            Font font1 = range1.Font;
            Font font2 = range2.Font;
            if (checkEqualGlow(font1, font2)
                && checkEqualReflection(font1, font2)
                && checkEqualShadow(font1, font2)
                && checkEqualColorEffect(font1, font2)
                )
            {
                return true;
            }
            else
            {
                return false;
            }
        }
        // hieu ung ruc ro
        public bool checkEqualGlow(Font font1, Font font2)
        {
            if (font1.Glow.Color.ObjectThemeColor.ToString() == font2.Glow.Color.ObjectThemeColor.ToString()
                && font1.Glow.Radius == font2.Glow.Radius
                && font1.Glow.Transparency == font2.Glow.Transparency)
            {
                return true;
            }
            else
            {
                return false;
            }
        }
        //hieu ung phan chieu
        public bool checkEqualReflection(Font font1, Font font2)
        {
            if (font1.Reflection.Blur == font2.Reflection.Blur
                && font1.Reflection.Size == font2.Reflection.Size
                && font1.Reflection.Transparency == font2.Reflection.Transparency
                && font1.Reflection.Offset == font2.Reflection.Offset)
            {
                return true;
            }
            else
            {
                return false;
            }
        }
        //hieu ung mau
        public bool checkEqualColorEffect(Font font1, Font font2)
        {
            if (font1.StylisticSet.ToString() == font2.StylisticSet.ToString()
                && font1.Outline == font2.Outline
                && font1.Ligatures.ToString() == font2.Ligatures.ToString()
                && font1.TextColor.ObjectThemeColor == font2.TextColor.ObjectThemeColor
                && font1.Fill.ForeColor.RGB == font2.Fill.ForeColor.RGB
                && font1.Fill.BackColor.RGB == font2.Fill.BackColor.RGB
                && font1.Line.Weight == font2.Line.Weight
                && font1.Line.DashStyle == font2.Line.DashStyle
                && font1.Line.ForeColor.RGB == font2.Line.ForeColor.RGB
                && font1.Line.BackColor.RGB == font2.Line.BackColor.RGB
                //&& font1.Line.ForeColor.Brightness == font2.Line.ForeColor.Brightness
                )
            {
                try
                {
                    if (font1.Line.ForeColor.Brightness == font2.Line.ForeColor.Brightness)
                    {
                        return true;
                    }
                    else
                    {
                        return false;
                    }
                }
                catch (Exception e)
                {
                    return true;
                }
            }
            else
            {
                return false;
            }
        }
        //hieu ung bong
        public bool checkEqualShadow(Font font1, Font font2)
        {
            if (font1.Shadow == font2.Shadow
                && font1.TextShadow.Blur == font2.TextShadow.Blur
                && font1.TextShadow.ForeColor.RGB == font2.TextShadow.ForeColor.RGB
                && font1.TextShadow.Obscured == font2.TextShadow.Obscured
                && font1.TextShadow.OffsetX == font2.TextShadow.OffsetX
                && font1.TextShadow.OffsetY == font2.TextShadow.OffsetY
                && font1.TextShadow.Size == font2.TextShadow.Size
                && font1.TextShadow.Transparency == font2.TextShadow.Transparency
                && font1.TextShadow.Type == font2.TextShadow.Type
                && font1.TextShadow.Style == font2.TextShadow.Style
                )
            {
                return true;
            }
            else
            {
                return false;
            }
        }
    }
}

