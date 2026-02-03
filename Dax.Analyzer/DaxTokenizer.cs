using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using TabularEditor.TOMWrapper;
using TabularEditor.TOMWrapper.Utils;

using DaxToken = TabularEditor.TOMWrapper.Utils.DaxToken;

public static class DaxTokenizer
{
    private static readonly MethodInfo _tokenizeString = FindTokenizeString();

    public static IList<DaxToken> Tokenize(IDaxDependantObject daxObject)
        => DaxDependencyHelper.Tokenize(daxObject);

    public static IList<DaxToken> Tokenize(string dax)
    {
        if (_tokenizeString == null)
            throw new MissingMethodException("No Tokenize(string) method found in TOMWrapper runtime.");

        var res = _tokenizeString.Invoke(null, new object[] { dax });

        if (res is IList<DaxToken> typed) return typed;
        if (res is IEnumerable enumerable) return enumerable.Cast<DaxToken>().ToList();

        throw new InvalidCastException($"Unexpected Tokenize(string) return type: {res?.GetType().FullName ?? "<null>"}");
    }

    private static MethodInfo FindTokenizeString()
    {
        var asm = typeof(DaxDependencyHelper).Assembly;

        foreach (var t in SafeGetTypes(asm))
        {
            var mi = t.GetMethods(BindingFlags.Public | BindingFlags.NonPublic | BindingFlags.Static)
                      .FirstOrDefault(m =>
                          m.Name == "Tokenize" &&
                          m.GetParameters().Length == 1 &&
                          m.GetParameters()[0].ParameterType == typeof(string));

            if (mi != null) return mi;
        }

        return null;
    }

    private static IEnumerable<Type> SafeGetTypes(Assembly asm)
    {
        try { return asm.GetTypes(); }
        catch (ReflectionTypeLoadException ex) { return ex.Types.Where(x => x != null); }
    }
}
