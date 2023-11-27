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
            Assert.That(ex, Is.Not.Null);
        }

        [Test]
        public void MessageConstructorTest()
        {
            const string message = "Exception message.";

            var ex = new ExcelMapperConvertException(message);

            Assert.That(ex.Message, Is.EqualTo(message));
        }

        [Test]
        public void MessageAndInnerExceptionConstructorTest()
        {
            const string message = "Exception message.";
            var baseEx = new StackOverflowException();

            var ex = new ExcelMapperConvertException(message, baseEx);

            Assert.That(ex.Message, Is.EqualTo(message));
            Assert.That(ex.InnerException?.GetType(), Is.EqualTo(baseEx.GetType()));
        }

        [Test]
        public void ParamsConstructorTest()
        {
            const string value = "test";
            var targetType = typeof(int);
            const int line = 1;
            const int column = 2;

            var ex = new ExcelMapperConvertException(value, targetType, line, column);

            Assert.That(ex.CellValue, Is.EqualTo(value));
            Assert.That(ex.TargetType, Is.EqualTo(targetType));
            Assert.That(ex.Line, Is.EqualTo(line));
            Assert.That(ex.Column, Is.EqualTo(column));
            Assert.That(ex.Message.Contains($"\"{value}\""));
            Assert.That(ex.Message.Contains(targetType.ToString()));
            Assert.That(ex.Message.Contains($"[L:{line}]"));
            Assert.That(ex.Message.Contains($"[C:{column}]"));
        }
    }
}
