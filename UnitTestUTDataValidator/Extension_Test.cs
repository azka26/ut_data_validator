using NUnit.Framework;
using UTDataValidator;

namespace UnitTestProject1
{
    public class Extension_Test
    {
        [Test]
        public void Table_TestValid()
        {
            var value = "table: test";
            Assert.IsTrue(value.IsTableInfo());
            Assert.AreEqual("test", value.GetTableName());
            
            value = "table : test";
            Assert.IsTrue(value.IsTableInfo());
            Assert.AreEqual("test", value.GetTableName());
            
            value = "Table :test";
            Assert.IsTrue(value.IsTableInfo());
            Assert.AreEqual("test", value.GetTableName());
        }
        
        [Test]
        public void Tabel_TestValid()
        {
            var value = "tabel: test";
            Assert.IsTrue(value.IsTableInfo());
            Assert.AreEqual("test", value.GetTableName());
            
            value = "tabel : test";
            Assert.IsTrue(value.IsTableInfo());
            Assert.AreEqual("test", value.GetTableName());
            
            value = "Tabel :test";
            Assert.IsTrue(value.IsTableInfo());
            Assert.AreEqual("test", value.GetTableName());
        }
        
        [Test]
        public void Invalid_Tabel_TestValid()
        {
            var value = "tabel: test test";
            Assert.IsFalse(value.IsTableInfo());
        }
        
        [Test]
        public void Invalid_Table_TestValid()
        {
            var value = "table: test test";
            Assert.IsFalse(value.IsTableInfo());
        }
    }
}