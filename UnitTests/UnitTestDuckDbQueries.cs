using System;
using ExcelDna.Integration;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using xlDuckDb;

namespace UnitTests
{
    [TestClass]
    public class UnitTestDuckDbQueries
    {
        [TestMethod]
        public void TestHttpParquetRead()
        {
            var data = DuckDbHelper.ExecuteQuery("SELECT * FROM 'https://duckdb.org/data/holdings.parquet'");
            Assert.IsNotNull(data);
        }

        [TestMethod]
        public void TestBigInt()
        {
            var data = DuckDbHelper.ExecuteQuery("SELECT 1::BIGINT");
            Assert.IsNotNull(data);
            Assert.IsInstanceOfType(data[1, 0], typeof(long));
            Assert.AreEqual(1L, data[1, 0]);
        }

        [TestMethod]
        public void TestNullBigInt()
        {
            var data = DuckDbHelper.ExecuteQuery("SELECT NULL::BIGINT");
            Assert.IsNotNull(data);
            Assert.IsInstanceOfType(data[1, 0], typeof(ExcelError));
            Assert.AreEqual(ExcelError.ExcelErrorNA, data[1, 0]);
        }

        [TestMethod]
        public void TestUBigInt()
        {
            var data = DuckDbHelper.ExecuteQuery("SELECT 1::UBIGINT");
            Assert.IsNotNull(data);
            Assert.IsInstanceOfType(data[1, 0], typeof(long));
            Assert.AreEqual(1L, data[1, 0]);
        }

        [TestMethod]
        public void TestBit()
        {
            var data = DuckDbHelper.ExecuteQuery("SELECT bitstring('0101011', 12)");
            Assert.IsNotNull(data);
            Assert.IsInstanceOfType(data[1, 0], typeof(string));
            Assert.AreEqual("000000101011", data[1, 0]);
        }

        [TestMethod]
        public void TestBlob()
        {
            var data = DuckDbHelper.ExecuteQuery("SELECT encode('my_string_with_ü')");
            Assert.IsNotNull(data);
            Assert.IsInstanceOfType(data[1, 0], typeof(string));
            Assert.AreEqual("my_string_with_ü", data[1, 0]);
        }

        [TestMethod]
        public void TestBoolean()
        {
            var data = DuckDbHelper.ExecuteQuery("SELECT true, false");
            Assert.IsNotNull(data);
            Assert.IsInstanceOfType(data[1, 0], typeof(bool));
            Assert.IsInstanceOfType(data[1, 1], typeof(bool));
            Assert.AreEqual(true, data[1, 0]);
            Assert.AreEqual(false, data[1, 1]);
        }

        [TestMethod]
        public void TestDate()
        {
            var data = DuckDbHelper.ExecuteQuery("SELECT make_date(1992, 9, 20)");
            Assert.IsNotNull(data);
            Assert.IsInstanceOfType(data[1, 0], typeof(DateTime));
            Assert.AreEqual(new DateTime(1992, 9, 20), data[1, 0]);
        }

        [TestMethod]
        public void TestDecimal()
        {
            var data = DuckDbHelper.ExecuteQuery("SELECT 123.45678::DECIMAL(8, 2)");
            Assert.IsNotNull(data);
            Assert.IsInstanceOfType(data[1, 0], typeof(double));
            Assert.AreEqual(123.46, data[1, 0]);
        }

        [TestMethod]
        public void TestDouble()
        {
            var data = DuckDbHelper.ExecuteQuery("SELECT 123.45678::DOUBLE");
            Assert.IsNotNull(data);
            Assert.IsInstanceOfType(data[1, 0], typeof(double));
            Assert.AreEqual(123.45678, data[1, 0]);
        }

        [TestMethod]
        public void TestNanDouble()
        {
            var data = DuckDbHelper.ExecuteQuery("SELECT 1.0/0::DOUBLE");
            Assert.IsNotNull(data);
            Assert.IsInstanceOfType(data[1, 0], typeof(ExcelError));
            Assert.AreEqual(ExcelError.ExcelErrorNum, data[1, 0]);
        }

        [TestMethod]
        public void TestFloat()
        {
            var data = DuckDbHelper.ExecuteQuery("SELECT 123.45678::FLOAT");
            Assert.IsNotNull(data);
            Assert.IsInstanceOfType(data[1, 0], typeof(double));
            Assert.IsTrue(123.45677947998 - (double) data[1, 0] < 1e-5);
        }

        [TestMethod]
        public void TestHugeInt()
        {
            var data = DuckDbHelper.ExecuteQuery("SELECT 1::HUGEINT");
            Assert.IsNotNull(data);
            Assert.IsInstanceOfType(data[1, 0], typeof(long));
            Assert.AreEqual(1L, data[1, 0]);
        }

        [TestMethod]
        public void TestUHugeInt()
        {
            var data = DuckDbHelper.ExecuteQuery("SELECT 1::UHUGEINT");
            Assert.IsNotNull(data);
            Assert.IsInstanceOfType(data[1, 0], typeof(long));
            Assert.AreEqual(1L, data[1, 0]);
        }


        [TestMethod]
        public void TestInteger()
        {
            var data = DuckDbHelper.ExecuteQuery("SELECT 1::INT");
            Assert.IsNotNull(data);
            Assert.IsInstanceOfType(data[1, 0], typeof(int));
            Assert.AreEqual(1, data[1, 0]);
        }

        [TestMethod]
        public void TestUInteger()
        {
            var data = DuckDbHelper.ExecuteQuery("SELECT 1::UINTEGER");
            Assert.IsNotNull(data);
            Assert.IsInstanceOfType(data[1, 0], typeof(uint));
            Assert.AreEqual((uint) 1, data[1, 0]);
        }

        [TestMethod]
        public void TestIntervalFromTime()
        {
            var now = DateTime.Now;
            var data = DuckDbHelper.ExecuteQuery("SELECT current_localtime()");
            Assert.IsNotNull(data);
            Assert.IsInstanceOfType(data[1, 0], typeof(DateTime));
            Assert.AreEqual(now.Hour, ((DateTime) data[1, 0]).Hour);
            Assert.AreEqual(now.Minute, ((DateTime) data[1, 0]).Minute);
            Assert.AreEqual(now.Second, ((DateTime) data[1, 0]).Second);
        }

        [TestMethod]
        public void TestIntervalFromFunc()
        {
            var data = DuckDbHelper.ExecuteQuery("SELECT to_days(10)");
            Assert.IsNotNull(data);
            Assert.IsInstanceOfType(data[1, 0], typeof(DateTime));
            Assert.IsTrue(Math.Abs(((DateTime) data[1, 0] - new DateTime(1899, 12, 30)).TotalDays - 10) < 0.01);
        }

        [TestMethod]
        public void TestJson()
        {
            var data = DuckDbHelper.ExecuteQuery("SELECT '{\"duck\": 42}'::JSON");
            Assert.IsNotNull(data);
            Assert.IsInstanceOfType(data[1, 0], typeof(string));
            Assert.AreEqual("{\"duck\": 42}", data[1, 0]);
        }

        [TestMethod]
        public void TestSmallInteger()
        {
            var data = DuckDbHelper.ExecuteQuery("SELECT 1::SMALLINT");
            Assert.IsNotNull(data);
            Assert.IsInstanceOfType(data[1, 0], typeof(short));
            Assert.AreEqual((short) 1, data[1, 0]);
        }

        [TestMethod]
        public void TestUSmallInteger()
        {
            var data = DuckDbHelper.ExecuteQuery("SELECT 1::USMALLINT");
            Assert.IsNotNull(data);
            Assert.IsInstanceOfType(data[1, 0], typeof(ushort));
            Assert.AreEqual((ushort) 1, data[1, 0]);
        }

        [TestMethod]
        public void TestTime()
        {
            var data = DuckDbHelper.ExecuteQuery("SELECT TIME '1992-09-20 11:30:00.123456'");
            Assert.IsNotNull(data);
            Assert.IsInstanceOfType(data[1, 0], typeof(DateTime));
            Assert.IsTrue(((DateTime) data[1, 0] - new DateTime(1899, 12, 30)).Hours == 11);
            Assert.IsTrue(((DateTime) data[1, 0] - new DateTime(1899, 12, 30)).Minutes == 30);
        }

        [TestMethod]
        public void TestTimeTz()
        {
            var data = DuckDbHelper.ExecuteQuery("SELECT TIMETZ '1992-09-20 11:30:00.123456+05:30'");
            Assert.IsNotNull(data);
            Assert.IsInstanceOfType(data[1, 0], typeof(DateTime));
            Assert.IsTrue(((DateTime) data[1, 0] - new DateTime(1899, 12, 30)).Hours == 6);
            Assert.IsTrue(((DateTime) data[1, 0] - new DateTime(1899, 12, 30)).Minutes == 0);
        }

        [TestMethod]
        public void TestTimeStamp()
        {
            var data = DuckDbHelper.ExecuteQuery("SELECT TIMESTAMP '1992-09-20 11:30:00';");
            Assert.IsNotNull(data);
            Assert.IsInstanceOfType(data[1, 0], typeof(DateTime));
            Assert.AreEqual(new DateTime(1992, 9, 20, 11, 30, 0), data[1, 0]);
        }

        [TestMethod]
        public void TestTimeStampTz()
        {
            var data = DuckDbHelper.ExecuteQuery("SELECT TIMESTAMPTZ '1992-09-20 11:30:00+01:00';");
            Assert.IsNotNull(data);
            Assert.IsInstanceOfType(data[1, 0], typeof(DateTime));
            Assert.AreEqual(new DateTime(1992, 9, 20, 10, 30, 0), data[1, 0]);
        }

        [TestMethod]
        public void TestTinyInteger()
        {
            var data = DuckDbHelper.ExecuteQuery("SELECT 1::TINYINT");
            Assert.IsNotNull(data);
            Assert.IsInstanceOfType(data[1, 0], typeof(sbyte));
            Assert.AreEqual((sbyte) 1, data[1, 0]);
        }

        [TestMethod]
        public void TestUTinyInteger()
        {
            var data = DuckDbHelper.ExecuteQuery("SELECT 1::UTINYINT");
            Assert.IsNotNull(data);
            Assert.IsInstanceOfType(data[1, 0], typeof(byte));
            Assert.AreEqual((byte) 1, data[1, 0]);
        }

        [TestMethod]
        public void TestUuid()
        {
            var data = DuckDbHelper.ExecuteQuery("SELECT 'eeccb8c5-9943-b2bb-bb5e-222f4e14b687'::UUID");
            Assert.IsNotNull(data);
            Assert.IsInstanceOfType(data[1, 0], typeof(string));
            Assert.AreEqual("eeccb8c5-9943-b2bb-bb5e-222f4e14b687", data[1, 0]);
        }

        [TestMethod]
        public void TestVarChar()
        {
            var data = DuckDbHelper.ExecuteQuery("SELECT 'Hello World'::VARCHAR");
            Assert.IsNotNull(data);
            Assert.IsInstanceOfType(data[1, 0], typeof(string));
            Assert.AreEqual("Hello World", data[1, 0]);
        }

        [TestMethod]
        public void TestNullVarchar()
        {
            var data = DuckDbHelper.ExecuteQuery("SELECT NULL::VARCHAR");
            Assert.IsNotNull(data);
            Assert.IsInstanceOfType(data[1, 0], typeof(ExcelError));
            Assert.AreEqual(ExcelError.ExcelErrorNA, data[1, 0]);
        }
    }
}