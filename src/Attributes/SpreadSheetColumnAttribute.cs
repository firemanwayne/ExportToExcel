﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;

namespace Simple.ExportToExcel.Attributes;

[AttributeUsage(AttributeTargets.Class)]
public class SpreadSheetAttribute : Attribute
{
    private readonly IList<SpreadSheetColumnAttribute> columns = new List<SpreadSheetColumnAttribute>();

    public SpreadSheetAttribute(string Name, Type Type) : base()
    {
        this.Name = Name;
        this.Type = Type;

        GetColumnIndexes();
    }

    public string Name { get; }
    public Type Type { get; }

    public IEnumerable<SpreadSheetColumnAttribute> Columns { get => columns.OrderBy(c => c.ColumnIndex); }

    void GetColumnIndexes()
    {
        PropertyInfo[] props = Type.GetProperties();
        foreach (PropertyInfo p in props)
        {
            SpreadSheetColumnAttribute columnAttr = p.GetCustomAttribute<SpreadSheetColumnAttribute>();
            if (columnAttr != null)
            {
                columnAttr.SetProperty(p);

                columns.Add(columnAttr);
            }
        }
    }
}

[AttributeUsage(AttributeTargets.Property)]
public class SpreadSheetColumnAttribute : Attribute
{
    public SpreadSheetColumnAttribute(string Name, int ColumnIndex) : base()
    {
        this.Name = Name;
        this.ColumnIndex = ColumnIndex;
    }

    public string Name { get; }
    public int ColumnIndex { get; }
    public PropertyInfo Property { get; private set; }

    internal void SetProperty(PropertyInfo p) => Property = p;
}
