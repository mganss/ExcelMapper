using System;
using Ganss.Excel.Exceptions;
using NUnit.Framework;

namespace ExcelMapper.Tests.ExceptionsTests
{
    public class ExcelMapperConvertExceptionTests
    {
        [Test]
        public void EmptyConstructorTest()
        {
            var ex = new ExcelMapperConvertException();
            Assert.NotNull(ex);
        }

        [Test]
        public void MessageConstructorTest()
        {
            const string message = "Exception message.";

            var ex = new ExcelMapperConvertException(message);

            Assert.AreEqual(message, ex.Message);
        }

        [Test]
        public void MessageAndInnerExceptionConstructorTest()
        {
            const string message = "Exception message.";
            var baseEx = new StackOverflowException();

            var ex = new ExcelMapperConvertException(message, baseEx);

            Assert.AreEqual(message, ex.Message);
            Assert.AreEqual(baseEx.GetType(), ex.InnerException?.GetType());
        }

        [Test]
        public void ParamsConstructorTest()
        {
            const string value = "test";
            var targetType = typeof(int);
            const int line = 1;
            const int column = 2;

            var ex = new ExcelMapperConvertException(value, targetType, line, column);

            Assert.AreEqual(value, ex.CellValue);
            Assert.AreEqual(targetType, ex.TargetType);
            Assert.AreEqual(line, ex.Line);
            Assert.AreEqual(column, ex.Column);
            Assert.That(ex.Message.Contains($"\"{value}\""));
            Assert.That(ex.Message.Contains(targetType.ToString()));
            Assert.That(ex.Message.Contains($"[L:{line}]"));
            Assert.That(ex.Message.Contains($"[C:{column}]"));
        }
    }
}
