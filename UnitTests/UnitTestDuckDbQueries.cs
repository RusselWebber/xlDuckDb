using ExcelDna.Integration;
using System.Diagnostics.CodeAnalysis;
using System.Text.Json;
using xlDuckDb;
using Xunit;

namespace UnitTests;

[Experimental("DuckDBNET001")]
public class UnitTestDuckDbQueries
{
    [Fact]
    public void TestHttpParquetRead()
    {
        var data = DuckDbHelper.ExecuteQuery("SELECT * FROM 'https://duckdb.org/data/holdings.parquet'");
        Assert.NotNull(data);
    }

    [Fact]
    public void TestBigInt()
    {
        var data = DuckDbHelper.ExecuteQuery("SELECT 1::BIGINT");
        Assert.NotNull(data);
        Assert.IsType<long>(data[1, 0]);
        Assert.Equal(1L, data[1, 0]);
    }

    [Fact]
    public void TestNullBigInt()
    {
        var data = DuckDbHelper.ExecuteQuery("SELECT NULL::BIGINT");
        Assert.NotNull(data);
        Assert.IsType<ExcelError>(data[1, 0]);
        Assert.Equal(ExcelError.ExcelErrorNA, data[1, 0]);
    }

    [Fact]
    public void TestUBigInt()
    {
        var data = DuckDbHelper.ExecuteQuery("SELECT 1::UBIGINT");
        Assert.NotNull(data);
        Assert.IsType<long>(data[1, 0]);
        Assert.Equal(1L, data[1, 0]);
    }

    [Fact]
    public void TestBit()
    {
        var data = DuckDbHelper.ExecuteQuery("SELECT bitstring('0101011', 12)");
        Assert.NotNull(data);
        Assert.IsType<string>(data[1, 0]);
        Assert.Equal("000000101011", data[1, 0]);
    }

    [Fact]
    public void TestBlob()
    {
        var data = DuckDbHelper.ExecuteQuery("SELECT encode('my_string_with_ü')");
        Assert.NotNull(data);
        Assert.IsType<string>(data[1, 0]);
        Assert.Equal("my_string_with_ü", data[1, 0]);
    }

    [Fact]
    public void TestBoolean()
    {
        var data = DuckDbHelper.ExecuteQuery("SELECT true, false");
        Assert.NotNull(data);
        Assert.IsType<bool>(data[1, 0]);
        Assert.IsType<bool>(data[1, 1]);
        Assert.True((bool)data[1, 0]);
        Assert.False((bool)data[1, 1]);
    }

    [Fact]
    public void TestDate()
    {
        var data = DuckDbHelper.ExecuteQuery("SELECT make_date(1992, 9, 20)");
        Assert.NotNull(data);
        Assert.IsType<DateTime>(data[1, 0]);
        Assert.Equal(new DateTime(1992, 9, 20), data[1, 0]);
    }

    [Fact]
    public void TestDateCast()
    {
        var data = DuckDbHelper.ExecuteQuery("SELECT '2021-04-01'::DATE");
        Assert.NotNull(data);
        Assert.IsType<DateTime>(data[1, 0]);
        Assert.Equal(new DateTime(2021, 4, 1), data[1, 0]);
    }

    [Fact]
    public void TestDecimal()
    {
        var data = DuckDbHelper.ExecuteQuery("SELECT 123.45678::DECIMAL(8, 2)");
        Assert.NotNull(data);
        Assert.IsType<double>(data[1, 0]);
        Assert.Equal(123.46, data[1, 0]);
    }

    [Fact]
    public void TestDouble()
    {
        var data = DuckDbHelper.ExecuteQuery("SELECT 123.45678::DOUBLE");
        Assert.NotNull(data);
        Assert.IsType<double>(data[1, 0]);
        Assert.Equal(123.45678, data[1, 0]);
    }

    [Fact]
    public void TestNanDouble()
    {
        var data = DuckDbHelper.ExecuteQuery("SELECT 1.0/0::DOUBLE");
        Assert.NotNull(data);
        Assert.IsType<ExcelError>(data[1, 0]);
        Assert.Equal(ExcelError.ExcelErrorNum, data[1, 0]);
    }

    [Fact]
    public void TestFloat()
    {
        var data = DuckDbHelper.ExecuteQuery("SELECT 123.45678::FLOAT");
        Assert.NotNull(data);
        Assert.IsType<double>(data[1, 0]);
        Assert.True(Math.Abs(123.45677947998 - (double)data[1, 0]) < 1e-5);
    }

    [Fact]
    public void TestHugeInt()
    {
        var data = DuckDbHelper.ExecuteQuery("SELECT 1::HUGEINT");
        Assert.NotNull(data);
        Assert.IsType<long>(data[1, 0]);
        Assert.Equal(1L, data[1, 0]);
    }

    [Fact]
    public void TestUHugeInt()
    {
        var data = DuckDbHelper.ExecuteQuery("SELECT 1::UHUGEINT");
        Assert.NotNull(data);
        Assert.IsType<long>(data[1, 0]);
        Assert.Equal(1L, data[1, 0]);
    }

    [Fact]
    public void TestInteger()
    {
        var data = DuckDbHelper.ExecuteQuery("SELECT 1::INT");
        Assert.NotNull(data);
        Assert.IsType<int>(data[1, 0]);
        Assert.Equal(1, data[1, 0]);
    }

    [Fact]
    public void TestUInteger()
    {
        var data = DuckDbHelper.ExecuteQuery("SELECT 1::UINTEGER");
        Assert.NotNull(data);
        Assert.IsType<uint>(data[1, 0]);
        Assert.Equal((uint)1, data[1, 0]);
    }

    [Fact]
    public void TestIntervalFromTime()
    {
        var now = DateTime.Now;
        var data = DuckDbHelper.ExecuteQuery("SELECT current_localtime()");
        Assert.NotNull(data);
        Assert.IsType<TimeOnly>(data[1, 0]);
        Assert.Equal(now.Hour, ((TimeOnly)data[1, 0]).Hour);
        Assert.Equal(now.Minute, ((TimeOnly)data[1, 0]).Minute);
        Assert.Equal(now.Second, ((TimeOnly)data[1, 0]).Second);
    }

    [Fact]
    public void TestIntervalFromFunc()
    {
        var data = DuckDbHelper.ExecuteQuery("SELECT to_days(10)");
        Assert.NotNull(data);
        Assert.IsType<DateTime>(data[1, 0]);
        Assert.True(Math.Abs(((DateTime)data[1, 0] - new DateTime(1899, 12, 30)).TotalDays - 10) < 0.01);
    }

    [Fact]
    public void TestJson()
    {
        var data = DuckDbHelper.ExecuteQuery("SELECT '{\"duck\": 42}'::JSON");
        Assert.NotNull(data);
        Assert.IsType<string>(data[1, 0]);
        Assert.Equal("{\"duck\": 42}", data[1, 0]);
    }

    [Fact(Skip = "DuckDb.Net does not support variants yet")]
    public void TestVariant()
    {
        var data = DuckDbHelper.ExecuteQuery("SELECT 42::VARIANT");
        Assert.NotNull(data);
        Assert.IsType<int>(data[1, 0]);
        Assert.Equal(42, data[1, 0]);
    }

    [Fact]
    public void TestSmallInteger()
    {
        var data = DuckDbHelper.ExecuteQuery("SELECT 1::SMALLINT");
        Assert.NotNull(data);
        Assert.IsType<short>(data[1, 0]);
        Assert.Equal((short)1, data[1, 0]);
    }

    [Fact]
    public void TestUSmallInteger()
    {
        var data = DuckDbHelper.ExecuteQuery("SELECT 1::USMALLINT");
        Assert.NotNull(data);
        Assert.IsType<ushort>(data[1, 0]);
        Assert.Equal((ushort)1, data[1, 0]);
    }

    [Fact]
    public void TestTime()
    {
        var data = DuckDbHelper.ExecuteQuery("SELECT TIME '1992-09-20 11:30:00.123456'");
        Assert.NotNull(data);
        Assert.IsType<TimeOnly>(data[1, 0]);
        Assert.True(((TimeOnly)data[1, 0]).Hour == 11);
        Assert.True(((TimeOnly)data[1, 0]).Minute == 30);
    }

    [Fact]
    public void TestTimeTz()
    {
        var data = DuckDbHelper.ExecuteQuery("SELECT TIMETZ '1992-09-20 11:30:00.123456+05:30'");
        Assert.NotNull(data);
        Assert.IsType<DateTime>(data[1, 0]);
        Assert.True(((DateTime)data[1, 0] - new DateTime(1899, 12, 30)).Hours == 6);
        Assert.True(((DateTime)data[1, 0] - new DateTime(1899, 12, 30)).Minutes == 0);
    }

    [Fact]
    public void TestTimeStamp()
    {
        var data = DuckDbHelper.ExecuteQuery("SELECT TIMESTAMP '1992-09-20 11:30:00';");
        Assert.NotNull(data);
        Assert.IsType<DateTime>(data[1, 0]);
        Assert.Equal(new DateTime(1992, 9, 20, 11, 30, 0), data[1, 0]);
    }

    [Fact]
    public void TestTimeStampTz()
    {
        var data = DuckDbHelper.ExecuteQuery("SELECT TIMESTAMPTZ '1992-09-20 11:30:00+01:00';");
        Assert.NotNull(data);
        Assert.IsType<DateTime>(data[1, 0]);
        Assert.Equal(new DateTime(1992, 9, 20, 10, 30, 0), data[1, 0]);
    }

    [Fact]
    public void TestList()
    {
        var data = DuckDbHelper.ExecuteQuery("SELECT ['apple', 'banana', 'cherry']");
        Assert.NotNull(data);
        var json = Assert.IsType<string>(data[1, 0]);
        var list = JsonSerializer.Deserialize<List<string>>(json);
        Assert.NotNull(list);
        Assert.Equal(3, list.Count);
        Assert.Equal("apple", list[0]);
        Assert.Equal("banana", list[1]);
        Assert.Equal("cherry", list[2]);
    }

    [Fact]
    public void TestNullList()
    {
        var data = DuckDbHelper.ExecuteQuery("SELECT NULL::INTEGER[]");
        Assert.NotNull(data);
        Assert.IsType<ExcelError>(data[1, 0]);
        Assert.Equal(ExcelError.ExcelErrorNA, data[1, 0]);
    }

    [Fact]
    public void TestStruct()
    {
        var data = DuckDbHelper.ExecuteQuery("SELECT {'id': 101, 'status': 'active'}");
        Assert.NotNull(data);
        var json = Assert.IsType<string>(data[1, 0]);
        var dict = JsonSerializer.Deserialize<Dictionary<string, JsonElement>>(json);
        Assert.NotNull(dict);
        Assert.Equal(101, dict["id"].GetInt32());
        Assert.Equal("active", dict["status"].GetString());
    }

    [Fact]
    public void TestMap()
    {
        var data = DuckDbHelper.ExecuteQuery("SELECT map(['color', 'shape'], ['red', 'circle'])");
        Assert.NotNull(data);
        var json = Assert.IsType<string>(data[1, 0]);
        var dict = JsonSerializer.Deserialize<Dictionary<string, string>>(json);
        Assert.NotNull(dict);
        Assert.Equal("red", dict["color"]);
        Assert.Equal("circle", dict["shape"]);
    }
}