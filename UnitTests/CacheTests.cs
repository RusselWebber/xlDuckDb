using System;
using Xunit;
using xlDuckDb;

public class CacheTests
{
    [Fact]
    public void CanCacheAndRetrieveArray()
    {
        var arr = new object[,] { { 1, 2 }, { 3, 4 } };
        var key = "test-array";
        
        // Should not exist initially
        Xunit.Assert.Null(xlAddIn.GetCachedRangeData(key));
        
        // Add to cache
        var addMethod = typeof(xlAddIn).GetMethod("CacheRangeData", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static | System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.FlattenHierarchy);
        addMethod.Invoke(null, new object[] { key, arr });
        
        // Should retrieve
        var cached = xlAddIn.GetCachedRangeData(key);
        Xunit.Assert.NotNull(cached);
        Xunit.Assert.Equal(arr, cached);
        
        // Remove
        Xunit.Assert.True(xlAddIn.RemoveCachedRangeData(key));
        Xunit.Assert.Null(xlAddIn.GetCachedRangeData(key));
    }

    [Fact]
    public void RemoveCachedRangeData_ReturnsFalseIfNotFound()
    {
        Xunit.Assert.False(xlAddIn.RemoveCachedRangeData("nonexistent-key"));
    }

    [Fact]
    public void ThrowsOnUnsupportedType()
    {
        // Should throw if passed an unsupported type
#pragma warning disable DuckDBNET001
        Xunit.Assert.Throws<ArgumentException>(() =>
            xlAddIn.DuckDbQuery("SELECT 1", "", new object[] { 123 })
        );
#pragma warning restore DuckDBNET001
    }
}
