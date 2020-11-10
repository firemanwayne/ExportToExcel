using NPOI.SS.UserModel;
using System.Collections.Generic;

namespace ExportToExcel.Styles
{
    public class ExcelColors
    {
        public ExcelColors(IndexedColors c, string Name)
        {
            Id = c.Index;
            this.Name = Name;
            RGBValue = $"rgb({c.RGB[0]}, {c.RGB[1]}, {c.RGB[2]})";

            IndexedColor = c;
        }

        public short Id { get; }
        public string Name { get; }
        public string RGBValue { get; }

        public IndexedColors IndexedColor { get; }

        public static readonly IDictionary<short, ExcelColors> ColorCollection = new Dictionary<short, ExcelColors>
        {
            { IndexedColors.Aqua.Index, new ExcelColors(IndexedColors.Aqua, "Aqua") },
            { IndexedColors.Automatic.Index, new ExcelColors(IndexedColors.Automatic, "Automatic") },

            { IndexedColors.Black.Index, new ExcelColors(IndexedColors.Black, "Black") },
            { IndexedColors.Blue.Index, new ExcelColors(IndexedColors.Blue, "Blue") },
            { IndexedColors.BlueGrey.Index, new ExcelColors(IndexedColors.BlueGrey, "Blue/Grey") },
            { IndexedColors.BrightGreen.Index, new ExcelColors(IndexedColors.BrightGreen, "Bright Green") },
            { IndexedColors.Brown.Index, new ExcelColors(IndexedColors.Brown, "Brown") },

            { IndexedColors.Coral.Index, new ExcelColors(IndexedColors.Coral, "Coral") },
            { IndexedColors.CornflowerBlue.Index, new ExcelColors(IndexedColors.CornflowerBlue, "Corn flower Blue") },

            { IndexedColors.DarkBlue.Index, new ExcelColors(IndexedColors.DarkBlue, "Dark Blue") },
            { IndexedColors.DarkGreen.Index, new ExcelColors(IndexedColors.DarkGreen, "Dark Green") },
            { IndexedColors.DarkRed.Index, new ExcelColors(IndexedColors.DarkRed, "Dark Red") },
            { IndexedColors.DarkTeal.Index, new ExcelColors(IndexedColors.DarkTeal, "Dark Teal") },
            { IndexedColors.DarkYellow.Index, new ExcelColors(IndexedColors.DarkYellow, "Dark Yellow") },

            { IndexedColors.Gold.Index, new ExcelColors(IndexedColors.Gold, "Gold") },
            { IndexedColors.Green.Index, new ExcelColors(IndexedColors.Green, "Green") },
            { IndexedColors.Grey25Percent.Index, new ExcelColors(IndexedColors.Grey25Percent, "Grey 25%") },
            { IndexedColors.Grey40Percent.Index, new ExcelColors(IndexedColors.Grey40Percent, "Grey 40%") },
            { IndexedColors.Grey50Percent.Index, new ExcelColors(IndexedColors.Grey50Percent, "Grey 50%") },
            { IndexedColors.Grey80Percent.Index, new ExcelColors(IndexedColors.Grey80Percent, "Grey 80%") },

            { IndexedColors.Indigo.Index, new ExcelColors(IndexedColors.Indigo, "Indigo") },

            { IndexedColors.Lavender.Index, new ExcelColors(IndexedColors.Lavender, "Lavender") },
            { IndexedColors.LemonChiffon.Index, new ExcelColors(IndexedColors.LemonChiffon, "Lemon Chiffon") },
            { IndexedColors.LightBlue.Index, new ExcelColors(IndexedColors.LightBlue, "Light Blue") },
            { IndexedColors.LightCornflowerBlue.Index, new ExcelColors(IndexedColors.LightCornflowerBlue, "Light Corn flower Blue") },
            { IndexedColors.LightGreen.Index, new ExcelColors(IndexedColors.LightGreen, "Light Green") },
            { IndexedColors.LightOrange.Index, new ExcelColors(IndexedColors.LightOrange, "Light Orange") },
            { IndexedColors.LightTurquoise.Index, new ExcelColors(IndexedColors.LightTurquoise, "Light Turquoise") },
            { IndexedColors.LightYellow.Index, new ExcelColors(IndexedColors.LightYellow, "Light Yellow") },

            { IndexedColors.Maroon.Index, new ExcelColors(IndexedColors.Maroon, "Maroon") },

            { IndexedColors.OliveGreen.Index, new ExcelColors(IndexedColors.OliveGreen, "Olive Green") },
            { IndexedColors.Orange.Index, new ExcelColors(IndexedColors.Orange, "Orange") },
            { IndexedColors.Orchid.Index, new ExcelColors(IndexedColors.Orchid, "Orchid") },

            { IndexedColors.PaleBlue.Index, new ExcelColors(IndexedColors.PaleBlue, "Pale Blue") },
            { IndexedColors.Pink.Index, new ExcelColors(IndexedColors.Pink, "Pink") },
            { IndexedColors.Plum.Index, new ExcelColors(IndexedColors.Plum, "Plum") },

            { IndexedColors.Red.Index, new ExcelColors(IndexedColors.Red, "Red") },
            { IndexedColors.Rose.Index, new ExcelColors(IndexedColors.Rose, "Rose") },
            { IndexedColors.RoyalBlue.Index, new ExcelColors(IndexedColors.RoyalBlue, "Royal Blue") },

            { IndexedColors.SeaGreen.Index, new ExcelColors(IndexedColors.SeaGreen, "Sea Green") },
            { IndexedColors.SkyBlue.Index, new ExcelColors(IndexedColors.SkyBlue, "Sky Blue") },

            { IndexedColors.Tan.Index, new ExcelColors(IndexedColors.Tan, "Tan") },
            { IndexedColors.Teal.Index, new ExcelColors(IndexedColors.Teal, "Teal") },
            { IndexedColors.Turquoise.Index, new ExcelColors(IndexedColors.Turquoise, "Turquoise") },

            { IndexedColors.Violet.Index, new ExcelColors(IndexedColors.Violet, "Violet") },

            { IndexedColors.White.Index, new ExcelColors(IndexedColors.White, "White") },

            { IndexedColors.Yellow.Index, new ExcelColors(IndexedColors.Yellow, "Yellow") },
        };        
    }
}
