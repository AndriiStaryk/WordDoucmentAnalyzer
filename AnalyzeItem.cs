using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DocumentAnalyzer;

internal struct AnalyzeItem
{
    public Margin? Margin { get; set; }
    public FontInfo? FontInfo { get; set; }
    public bool? Pictures {  get; set; }
    public bool? Tables {  get; set; }
}


internal struct FontInfo
{
    public string Name { get; set; }
    public double Size { get; set; }
    public override string ToString() => $"Font: {Name}, Size: {Size}";
}

internal struct Margin
{
    //stored in twips
    public UInt32 Left { get; set; }
    public UInt32 Top { get; set; }
    public UInt32 Right { get; set; }
    public UInt32 Bottom { get; set; }

    public Margin(double topInches, double bottomInches, double leftInches, double rightInches)
    {
        // 1 twip = 1/1440 inch
        const double coefficient = 1440;

        Top = (UInt32)(topInches * coefficient);
        Bottom = (UInt32)(bottomInches * coefficient);
        Left = (UInt32)(leftInches * coefficient);
        Right = (UInt32)(rightInches * coefficient);
    }

    public Margin(UInt32 topTwips, UInt32 bottomTwips, UInt32 leftTwips, UInt32 rightTwips)
    {
        Top = topTwips;
        Bottom = bottomTwips;
        Left = leftTwips;
        Right = rightTwips;
    }

}

