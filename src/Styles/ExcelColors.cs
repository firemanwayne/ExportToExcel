using NPOI.SS.UserModel;
using System.Collections.Generic;

namespace Simple.ExportToExcel.Styles
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
            { IndexedColors.Aqua.Index, new(IndexedColors.Aqua, "Aqua") },
            { IndexedColors.Automatic.Index, new(IndexedColors.Automatic, "Automatic") },

            { IndexedColors.Black.Index, new(IndexedColors.Black, "Black") },
            { IndexedColors.Blue.Index, new(IndexedColors.Blue, "Blue") },
            { IndexedColors.BlueGrey.Index, new(IndexedColors.BlueGrey, "Blue/Grey") },
            { IndexedColors.BrightGreen.Index, new(IndexedColors.BrightGreen, "Bright Green") },
            { IndexedColors.Brown.Index, new(IndexedColors.Brown, "Brown") },

            { IndexedColors.Coral.Index, new(IndexedColors.Coral, "Coral") },
            { IndexedColors.CornflowerBlue.Index, new(IndexedColors.CornflowerBlue, "Corn flower Blue") },

            { IndexedColors.DarkBlue.Index, new(IndexedColors.DarkBlue, "Dark Blue") },
            { IndexedColors.DarkGreen.Index, new(IndexedColors.DarkGreen, "Dark Green") },
            { IndexedColors.DarkRed.Index, new(IndexedColors.DarkRed, "Dark Red") },
            { IndexedColors.DarkTeal.Index, new(IndexedColors.DarkTeal, "Dark Teal") },
            { IndexedColors.DarkYellow.Index, new(IndexedColors.DarkYellow, "Dark Yellow") },

            { IndexedColors.Gold.Index, new(IndexedColors.Gold, "Gold") },
            { IndexedColors.Green.Index, new(IndexedColors.Green, "Green") },
            { IndexedColors.Grey25Percent.Index, new(IndexedColors.Grey25Percent, "Grey 25%") },
            { IndexedColors.Grey40Percent.Index, new(IndexedColors.Grey40Percent, "Grey 40%") },
            { IndexedColors.Grey50Percent.Index, new(IndexedColors.Grey50Percent, "Grey 50%") },
            { IndexedColors.Grey80Percent.Index, new(IndexedColors.Grey80Percent, "Grey 80%") },

            { IndexedColors.Indigo.Index, new(IndexedColors.Indigo, "Indigo") },

            { IndexedColors.Lavender.Index, new(IndexedColors.Lavender, "Lavender") },
            { IndexedColors.LemonChiffon.Index, new(IndexedColors.LemonChiffon, "Lemon Chiffon") },
            { IndexedColors.LightBlue.Index, new(IndexedColors.LightBlue, "Light Blue") },
            { IndexedColors.LightCornflowerBlue.Index, new(IndexedColors.LightCornflowerBlue, "Light Corn flower Blue") },
            { IndexedColors.LightGreen.Index, new(IndexedColors.LightGreen, "Light Green") },
            { IndexedColors.LightOrange.Index, new(IndexedColors.LightOrange, "Light Orange") },
            { IndexedColors.LightTurquoise.Index, new(IndexedColors.LightTurquoise, "Light Turquoise") },
            { IndexedColors.LightYellow.Index, new(IndexedColors.LightYellow, "Light Yellow") },

            { IndexedColors.Maroon.Index, new(IndexedColors.Maroon, "Maroon") },

            { IndexedColors.OliveGreen.Index, new(IndexedColors.OliveGreen, "Olive Green") },
            { IndexedColors.Orange.Index, new(IndexedColors.Orange, "Orange") },
            { IndexedColors.Orchid.Index, new(IndexedColors.Orchid, "Orchid") },

            { IndexedColors.PaleBlue.Index, new(IndexedColors.PaleBlue, "Pale Blue") },
            { IndexedColors.Pink.Index, new(IndexedColors.Pink, "Pink") },
            { IndexedColors.Plum.Index, new(IndexedColors.Plum, "Plum") },

            { IndexedColors.Red.Index, new(IndexedColors.Red, "Red") },
            { IndexedColors.Rose.Index, new(IndexedColors.Rose, "Rose") },
            { IndexedColors.RoyalBlue.Index, new(IndexedColors.RoyalBlue, "Royal Blue") },

            { IndexedColors.SeaGreen.Index, new(IndexedColors.SeaGreen, "Sea Green") },
            { IndexedColors.SkyBlue.Index, new(IndexedColors.SkyBlue, "Sky Blue") },

            { IndexedColors.Tan.Index, new(IndexedColors.Tan, "Tan") },
            { IndexedColors.Teal.Index, new(IndexedColors.Teal, "Teal") },
            { IndexedColors.Turquoise.Index, new(IndexedColors.Turquoise, "Turquoise") },

            { IndexedColors.Violet.Index, new(IndexedColors.Violet, "Violet") },

            { IndexedColors.White.Index, new(IndexedColors.White, "White") },

            { IndexedColors.Yellow.Index, new(IndexedColors.Yellow, "Yellow") },
        };
    }
}
