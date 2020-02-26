using Ganss.Excel.Exceptions;
using NUnit.Framework;

namespace ExcelMapper.Tests.ExceptionsTests
{
    [TestFixture]
    public class IsPrimitiveTypeExceptionTest
    {
        private IsPrimitiveTypeException _isPrimitiveTypeException;

        [SetUp]
        public void InitializeTest()
        {
            _isPrimitiveTypeException = new IsPrimitiveTypeException(typeof(double).Name);
        }

        [Test]
        public void ThrowsWellFormatedExceptionMessage()
        {
            var message = "Double cannot be a mapping type because it is primitive";

            Assert.AreEqual(message, _isPrimitiveTypeException.Message);
        }
    }
}
