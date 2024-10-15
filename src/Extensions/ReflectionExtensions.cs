using System.Collections.Generic;

namespace System.Reflection;

public static class ReflectionExtensions
{
    public static bool IsList(this PropertyInfo p)
    {
        bool IsGeneric = p.PropertyType.IsGenericType;
        if (!IsGeneric)
        {
            return false;
        }

        bool IsGenericList = p.PropertyType.GetGenericTypeDefinition() == typeof(IList<>);
        if (!IsGenericList)
        {
            return false;
        }

        return p.PropertyType.GetGenericTypeDefinition() == typeof(IEnumerable<>);
    }
}
