namespace Simple.ExportToExcel;

using NPOI.SS.UserModel;

using System;

public class HorizontalAlignmentChangedEventArgs : EventArgs
{
    public HorizontalAlignment SelectedAlignment { get; }

    public HorizontalAlignmentChangedEventArgs(HorizontalAlignment args)
    {
        SelectedAlignment = args;
    }
}
