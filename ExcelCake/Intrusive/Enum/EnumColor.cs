using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace ExcelCake.Intrusive
{
    public enum EnumColor
    {
        //
        // 摘要:
        //     Gets a system-defined color that has an ARGB value of #FF66CDAA.
        MediumAquamarine,
        //
        // 摘要:
        //     Gets a system-defined color that has an ARGB value of #FF0000CD.
        MediumBlue,
        //
        // 摘要:
        //     Gets a system-defined color that has an ARGB value of #FFBA55D3.
        MediumOrchid,
        //
        // 摘要:
        //     Gets a system-defined color that has an ARGB value of #FF9370DB.
        MediumPurple,
        //
        // 摘要:
        //     Gets a system-defined color that has an ARGB value of #FF3CB371.
        MediumSeaGreen,
        //
        // 摘要:
        //     Gets a system-defined color that has an ARGB value of #FF7B68EE.
        MediumSlateBlue,
        //
        // 摘要:
        //     Gets a system-defined color that has an ARGB value of #FF00FA9A.
        MediumSpringGreen,
        //
        // 摘要:
        //     Gets a system-defined color that has an ARGB value of #FF800000.
        Maroon,
        //
        // 摘要:
        //     Gets a system-defined color that has an ARGB value of #FF48D1CC.
        MediumTurquoise,
        //
        // 摘要:
        //     Gets a system-defined color that has an ARGB value of #FF191970.
        MidnightBlue,
        //
        // 摘要:
        //     Gets a system-defined color that has an ARGB value of #FFF5FFFA.
        MintCream,
        //
        // 摘要:
        //     Gets a system-defined color that has an ARGB value of #FFFFE4E1.
        MistyRose,
        //
        // 摘要:
        //     Gets a system-defined color that has an ARGB value of #FFFFE4B5.
        Moccasin,
        //
        // 摘要:
        //     Gets a system-defined color that has an ARGB value of #FFFFDEAD.
        NavajoWhite,
        //
        // 摘要:
        //     Gets a system-defined color that has an ARGB value of #FF000080.
        Navy,
        //
        // 摘要:
        //     Gets a system-defined color that has an ARGB value of #FFFDF5E6.
        OldLace,
        //
        // 摘要:
        //     Gets a system-defined color that has an ARGB value of #FFC71585.
        MediumVioletRed,
        //
        // 摘要:
        //     Gets a system-defined color that has an ARGB value of #FFFF00FF.
        Magenta,
        //
        // 摘要:
        //     Gets a system-defined color that has an ARGB value of #FFFAF0E6.
        Linen,
        //
        // 摘要:
        //     Gets a system-defined color that has an ARGB value of #FF32CD32.
        LimeGreen,
        //
        // 摘要:
        //     Gets a system-defined color that has an ARGB value of #FFFFF0F5.
        LavenderBlush,
        //
        // 摘要:
        //     Gets a system-defined color that has an ARGB value of #FF7CFC00.
        LawnGreen,
        //
        // 摘要:
        //     Gets a system-defined color that has an ARGB value of #FFFFFACD.
        LemonChiffon,
        //
        // 摘要:
        //     Gets a system-defined color that has an ARGB value of #FFADD8E6.
        LightBlue,
        //
        // 摘要:
        //     Gets a system-defined color that has an ARGB value of #FFF08080.
        LightCoral,
        //
        // 摘要:
        //     Gets a system-defined color that has an ARGB value of #FFE0FFFF.
        LightCyan,
        //
        // 摘要:
        //     Gets a system-defined color that has an ARGB value of #FFFAFAD2.
        LightGoldenrodYellow,
        //
        // 摘要:
        //     Gets a system-defined color that has an ARGB value of #FFD3D3D3.
        LightGray,
        //
        // 摘要:
        //     Gets a system-defined color that has an ARGB value of #FF90EE90.
        LightGreen,
        //
        // 摘要:
        //     Gets a system-defined color that has an ARGB value of #FFFFB6C1.
        LightPink,
        //
        // 摘要:
        //     Gets a system-defined color that has an ARGB value of #FFFFA07A.
        LightSalmon,
        //
        // 摘要:
        //     Gets a system-defined color that has an ARGB value of #FF20B2AA.
        LightSeaGreen,
        //
        // 摘要:
        //     Gets a system-defined color that has an ARGB value of #FF87CEFA.
        LightSkyBlue,
        //
        // 摘要:
        //     Gets a system-defined color that has an ARGB value of #FF778899.
        LightSlateGray,
        //
        // 摘要:
        //     Gets a system-defined color that has an ARGB value of #FFB0C4DE.
        LightSteelBlue,
        //
        // 摘要:
        //     Gets a system-defined color that has an ARGB value of #FFFFFFE0.
        LightYellow,
        //
        // 摘要:
        //     Gets a system-defined color that has an ARGB value of #FF00FF00.
        Lime,
        //
        // 摘要:
        //     Gets a system-defined color that has an ARGB value of #FF808000.
        Olive,
        //
        // 摘要:
        //     Gets a system-defined color that has an ARGB value of #FF6B8E23.
        OliveDrab,
        //
        // 摘要:
        //     Gets a system-defined color that has an ARGB value of #FFFFA500.
        Orange,
        //
        // 摘要:
        //     Gets a system-defined color that has an ARGB value of #FFFF4500.
        OrangeRed,
        //
        // 摘要:
        //     Gets a system-defined color that has an ARGB value of #FFC0C0C0.
        Silver,
        //
        // 摘要:
        //     Gets a system-defined color that has an ARGB value of #FF87CEEB.
        SkyBlue,
        //
        // 摘要:
        //     Gets a system-defined color that has an ARGB value of #FF6A5ACD.
        SlateBlue,
        //
        // 摘要:
        //     Gets a system-defined color that has an ARGB value of #FF708090.
        SlateGray,
        //
        // 摘要:
        //     Gets a system-defined color that has an ARGB value of #FFFFFAFA.
        Snow,
        //
        // 摘要:
        //     Gets a system-defined color that has an ARGB value of #FF00FF7F.
        SpringGreen,
        //
        // 摘要:
        //     Gets a system-defined color that has an ARGB value of #FF4682B4.
        SteelBlue,
        //
        // 摘要:
        //     Gets a system-defined color that has an ARGB value of #FFD2B48C.
        Tan,
        //
        // 摘要:
        //     Gets a system-defined color that has an ARGB value of #FF008080.
        Teal,
        //
        // 摘要:
        //     Gets a system-defined color that has an ARGB value of #FFD8BFD8.
        Thistle,
        //
        // 摘要:
        //     Gets a system-defined color that has an ARGB value of #FFFF6347.
        Tomato,
        //
        // 摘要:
        //     Gets a system-defined color.
        Transparent,
        //
        // 摘要:
        //     Gets a system-defined color that has an ARGB value of #FF40E0D0.
        Turquoise,
        //
        // 摘要:
        //     Gets a system-defined color that has an ARGB value of #FFEE82EE.
        Violet,
        //
        // 摘要:
        //     Gets a system-defined color that has an ARGB value of #FFF5DEB3.
        Wheat,
        //
        // 摘要:
        //     Gets a system-defined color that has an ARGB value of #FFFFFFFF.
        White,
        //
        // 摘要:
        //     Gets a system-defined color that has an ARGB value of #FFF5F5F5.
        WhiteSmoke,
        //
        // 摘要:
        //     Gets a system-defined color that has an ARGB value of #FFA0522D.
        Sienna,
        //
        // 摘要:
        //     Gets a system-defined color that has an ARGB value of #FFE6E6FA.
        Lavender,
        //
        // 摘要:
        //     Gets a system-defined color that has an ARGB value of #FFFFF5EE.
        SeaShell,
        //
        // 摘要:
        //     Gets a system-defined color that has an ARGB value of #FFF4A460.
        SandyBrown,
        //
        // 摘要:
        //     Gets a system-defined color that has an ARGB value of #FFDA70D6.
        Orchid,
        //
        // 摘要:
        //     Gets a system-defined color that has an ARGB value of #FFEEE8AA.
        PaleGoldenrod,
        //
        // 摘要:
        //     Gets a system-defined color that has an ARGB value of #FF98FB98.
        PaleGreen,
        //
        // 摘要:
        //     Gets a system-defined color that has an ARGB value of #FFAFEEEE.
        PaleTurquoise,
        //
        // 摘要:
        //     Gets a system-defined color that has an ARGB value of #FFDB7093.
        PaleVioletRed,
        //
        // 摘要:
        //     Gets a system-defined color that has an ARGB value of #FFFFEFD5.
        PapayaWhip,
        //
        // 摘要:
        //     Gets a system-defined color that has an ARGB value of #FFFFDAB9.
        PeachPuff,
        //
        // 摘要:
        //     Gets a system-defined color that has an ARGB value of #FFCD853F.
        Peru,
        //
        // 摘要:
        //     Gets a system-defined color that has an ARGB value of #FFFFC0CB.
        Pink,
        //
        // 摘要:
        //     Gets a system-defined color that has an ARGB value of #FFDDA0DD.
        Plum,
        //
        // 摘要:
        //     Gets a system-defined color that has an ARGB value of #FFB0E0E6.
        PowderBlue,
        //
        // 摘要:
        //     Gets a system-defined color that has an ARGB value of #FF800080.
        Purple,
        //
        // 摘要:
        //     Gets a system-defined color that has an ARGB value of #FFFF0000.
        Red,
        //
        // 摘要:
        //     Gets a system-defined color that has an ARGB value of #FFBC8F8F.
        RosyBrown,
        //
        // 摘要:
        //     Gets a system-defined color that has an ARGB value of #FF4169E1.
        RoyalBlue,
        //
        // 摘要:
        //     Gets a system-defined color that has an ARGB value of #FF8B4513.
        SaddleBrown,
        //
        // 摘要:
        //     Gets a system-defined color that has an ARGB value of #FFFA8072.
        Salmon,
        //
        // 摘要:
        //     Gets a system-defined color that has an ARGB value of #FF2E8B57.
        SeaGreen,
        //
        // 摘要:
        //     Gets a system-defined color that has an ARGB value of #FFFFFF00.
        Yellow,
        //
        // 摘要:
        //     Gets a system-defined color that has an ARGB value of #FFF0E68C.
        Khaki,
        //
        // 摘要:
        //     Gets a system-defined color that has an ARGB value of #FF00FFFF.
        Cyan,
        //
        // 摘要:
        //     Gets a system-defined color that has an ARGB value of #FF8B008B.
        DarkMagenta,
        //
        // 摘要:
        //     Gets a system-defined color that has an ARGB value of #FFBDB76B.
        DarkKhaki,
        //
        // 摘要:
        //     Gets a system-defined color that has an ARGB value of #FF006400.
        DarkGreen,
        //
        // 摘要:
        //     Gets a system-defined color that has an ARGB value of #FFA9A9A9.
        DarkGray,
        //
        // 摘要:
        //     Gets a system-defined color that has an ARGB value of #FFB8860B.
        DarkGoldenrod,
        //
        // 摘要:
        //     Gets a system-defined color that has an ARGB value of #FF008B8B.
        DarkCyan,
        //
        // 摘要:
        //     Gets a system-defined color that has an ARGB value of #FF00008B.
        DarkBlue,
        //
        // 摘要:
        //     Gets a system-defined color that has an ARGB value of #FFFFFFF0.
        Ivory,
        //
        // 摘要:
        //     Gets a system-defined color that has an ARGB value of #FFDC143C.
        Crimson,
        //
        // 摘要:
        //     Gets a system-defined color that has an ARGB value of #FFFFF8DC.
        Cornsilk,
        //
        // 摘要:
        //     Gets a system-defined color that has an ARGB value of #FF6495ED.
        CornflowerBlue,
        //
        // 摘要:
        //     Gets a system-defined color that has an ARGB value of #FFFF7F50.
        Coral,
        //
        // 摘要:
        //     Gets a system-defined color that has an ARGB value of #FFD2691E.
        Chocolate,
        //
        // 摘要:
        //     Gets a system-defined color that has an ARGB value of #FF556B2F.
        DarkOliveGreen,
        //
        // 摘要:
        //     Gets a system-defined color that has an ARGB value of #FF7FFF00.
        Chartreuse,
        //
        // 摘要:
        //     Gets a system-defined color that has an ARGB value of #FFDEB887.
        BurlyWood,
        //
        // 摘要:
        //     Gets a system-defined color that has an ARGB value of #FFA52A2A.
        Brown,
        //
        // 摘要:
        //     Gets a system-defined color that has an ARGB value of #FF8A2BE2.
        BlueViolet,
        //
        // 摘要:
        //     Gets a system-defined color that has an ARGB value of #FF0000FF.
        Blue,
        //
        // 摘要:
        //     Gets a system-defined color that has an ARGB value of #FFFFEBCD.
        BlanchedAlmond,
        //
        // 摘要:
        //     Gets a system-defined color that has an ARGB value of #FF000000.
        Black,
        //
        // 摘要:
        //     Gets a system-defined color that has an ARGB value of #FFFFE4C4.
        Bisque,
        //
        // 摘要:
        //     Gets a system-defined color that has an ARGB value of #FFF5F5DC.
        Beige,
        //
        // 摘要:
        //     Gets a system-defined color that has an ARGB value of #FFF0FFFF.
        Azure,
        //
        // 摘要:
        //     Gets a system-defined color that has an ARGB value of #FF7FFFD4.
        Aquamarine,
        //
        // 摘要:
        //     Gets a system-defined color that has an ARGB value of #FF00FFFF.
        Aqua,
        //
        // 摘要:
        //     Gets a system-defined color that has an ARGB value of #FFFAEBD7.
        AntiqueWhite,
        //
        // 摘要:
        //     Gets a system-defined color that has an ARGB value of #FFF0F8FF.
        AliceBlue,
        //
        // 摘要:
        //     Gets a system-defined color that has an ARGB value of #FF5F9EA0.
        CadetBlue,
        //
        // 摘要:
        //     Gets a system-defined color that has an ARGB value of #FFFF8C00.
        DarkOrange,
        //
        // 摘要:
        //     Gets a system-defined color that has an ARGB value of #FF9ACD32.
        YellowGreen,
        //
        // 摘要:
        //     Gets a system-defined color that has an ARGB value of #FF8B0000.
        DarkRed,
        //
        // 摘要:
        //     Gets a system-defined color that has an ARGB value of #FF4B0082.
        Indigo,
        //
        // 摘要:
        //     Gets a system-defined color that has an ARGB value of #FFCD5C5C.
        IndianRed,
        //
        // 摘要:
        //     Gets a system-defined color that has an ARGB value of #FF9932CC.
        DarkOrchid,
        //
        // 摘要:
        //     Gets a system-defined color that has an ARGB value of #FFF0FFF0.
        Honeydew,
        //
        // 摘要:
        //     Gets a system-defined color that has an ARGB value of #FFADFF2F.
        GreenYellow,
        //
        // 摘要:
        //     Gets a system-defined color that has an ARGB value of #FF008000.
        Green,
        //
        // 摘要:
        //     Gets a system-defined color that has an ARGB value of #FF808080.
        Gray,
        //
        // 摘要:
        //     Gets a system-defined color that has an ARGB value of #FFDAA520.
        Goldenrod,
        //
        // 摘要:
        //     Gets a system-defined color that has an ARGB value of #FFFFD700.
        Gold,
        //
        // 摘要:
        //     Gets a system-defined color that has an ARGB value of #FFF8F8FF.
        GhostWhite,
        //
        // 摘要:
        //     Gets a system-defined color that has an ARGB value of #FFDCDCDC.
        Gainsboro,
        //
        // 摘要:
        //     Gets a system-defined color that has an ARGB value of #FFFF00FF.
        Fuchsia,
        //
        // 摘要:
        //     Gets a system-defined color that has an ARGB value of #FF228B22.
        ForestGreen,
        //
        // 摘要:
        //     Gets a system-defined color that has an ARGB value of #FFFF69B4.
        HotPink,
        //
        // 摘要:
        //     Gets a system-defined color that has an ARGB value of #FFB22222.
        Firebrick,
        //
        // 摘要:
        //     Gets a system-defined color that has an ARGB value of #FFFFFAF0.
        FloralWhite,
        //
        // 摘要:
        //     Gets a system-defined color that has an ARGB value of #FF1E90FF.
        DodgerBlue,
        //
        // 摘要:
        //     Gets a system-defined color that has an ARGB value of #FF696969.
        DimGray,
        //
        // 摘要:
        //     Gets a system-defined color that has an ARGB value of #FF00BFFF.
        DeepSkyBlue,
        //
        // 摘要:
        //     Gets a system-defined color that has an ARGB value of #FFFF1493.
        DeepPink,
        //
        // 摘要:
        //     Gets a system-defined color that has an ARGB value of #FF9400D3.
        DarkViolet,
        //
        // 摘要:
        //     Gets a system-defined color that has an ARGB value of #FF00CED1.
        DarkTurquoise,
        //
        // 摘要:
        //     Gets a system-defined color that has an ARGB value of #FF2F4F4F.
        DarkSlateGray,
        //
        // 摘要:
        //     Gets a system-defined color that has an ARGB value of #FF483D8B.
        DarkSlateBlue,
        //
        // 摘要:
        //     Gets a system-defined color that has an ARGB value of #FF8FBC8F.
        DarkSeaGreen,
        //
        // 摘要:
        //     Gets a system-defined color that has an ARGB value of #FFE9967A.
        DarkSalmon,
    }
}
