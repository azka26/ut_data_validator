using Microsoft.VisualStudio.TestTools.UnitTesting;
using System;
using System.IO;
using UTDataValidator;
using System.Data.Common;

namespace UnitTestProject1
{
    public class UnitTestBase
    {
        protected ExcelValidator GetExcelValidator(string excelPath, string worksheetInitData, string worksheetExpected)
        {
            IEventExcelValidator eventExcelValidator = new SampleEventValidator();
            ExcelValidator excelValidator = new ExcelValidator(
                new Assertion()
                , excelPath: excelPath
                , worksheetInitData: "Sheet1"
                , worksheetExpectedData: "Sheet2",
                eventExcelValidator
            );

            excelValidator.ExecuteAction();
            return excelValidator;
        }
    }

    [TestClass]
    public class OutboundUnitTest : UnitTestBase
    {
        [TestMethod]
        public void Picking_Test()
        {
            string excelPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "excel", "sample.xlsx");
            ExcelValidator excelValidator = GetExcelValidator(excelPath, "Sheet1", "Sheet2");
            excelValidator.Validate();
        }
    }
}
