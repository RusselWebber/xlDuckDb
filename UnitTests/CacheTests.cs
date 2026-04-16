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
        Assert.Null(xlAddIn.GetCachedRangeData(key));
        
        // Add to cache
        var addMethod = typeof(xlAddIn).GetMethod("CacheRangeData", System.Reflection.BindingFlags.NonPublic | System.Reflection.BindingFlags.Static | System.Reflection.BindingFlags.Public | System.Reflection.BindingFlags.FlattenHierarchy);
        Assert.NotNull(addMethod); // Ensure method exists
        addMethod.Invoke(null, new object[] { key, arr });
        
        // Should retrieve
        var cached = xlAddIn.GetCachedRangeData(key);
        Assert.NotNull(cached);
        Assert.Equal(arr, cached!); // ! because NotNull above
        
        // Remove
        Assert.True(xlAddIn.RemoveCachedRangeData(key));
        Assert.Null(xlAddIn.GetCachedRangeData(key));
    }

    [Fact]
    public void RemoveCachedRangeData_ReturnsFalseIfNotFound()
    {
        Assert.False(xlAddIn.RemoveCachedRangeData("nonexistent-key"));
    }

    [Fact]
    public void ThrowsOnUnsupportedType()
    {
        // Should throw if passed an unsupported type
#pragma warning disable DuckDBNET001
        Assert.Throws<ArgumentException>(() =>
            xlAddIn.DuckDbQuery("SELECT 1", "", new object[] { 123 })
        );
#pragma warning restore DuckDBNET001
    }
}
