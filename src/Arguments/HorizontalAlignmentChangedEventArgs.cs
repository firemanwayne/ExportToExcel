using NPOI.SS.UserModel;
using System;

namespace Simple.ExportToExcel
{
    public class HorizontalAlignmentChangedEventArgs : EventArgs
    {
        public HorizontalAlignment SelectedAlignment { get; }

        public HorizontalAlignmentChangedEventArgs(HorizontalAlignment args)
        {
            SelectedAlignment = args;
        }
    }
}
