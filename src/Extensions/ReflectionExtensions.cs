namespace System.Reflection;

/// <summary>
/// Extension methods for <see cref="PropertyInfo"/> to simplify reflection-based type checks.
/// </summary>
public static class ReflectionExtensions
{
    /// <summary>
    /// Determines whether the property's type is a generic <see cref="IList{T}"/>.
    /// </summary>
    /// <param name="p">The property to inspect.</param>
    /// <returns><c>true</c> if the property type is <see cref="IList{T}"/>; otherwise <c>false</c>.</returns>
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
