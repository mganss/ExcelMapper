using NPOI.SS.UserModel;
using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Text;
using Ganss.Excel.Exceptions;
using System.Runtime.Serialization.Formatters.Binary;
using System.IO;

namespace Ganss.Excel.Tests
{
    public class Tests
    {
        [SetUp]
        public void Setup()
        {
            Environment.CurrentDirectory = TestContext.CurrentContext.TestDirectory;
        }

        private class Product
        {
            public string Name { get; set; }
            [Column("Number")]
            public int NumberInStock { get; set; }
            public decimal Price { get; set; }
            public string Value { get; set; }

            public override bool Equals(object obj)
            {
                if (!(obj is Product o)) return false;
                return o.Name == Name && o.NumberInStock == NumberInStock && o.Price == Price && o.Value == Value;
            }

            public override int GetHashCode()
            {
                return (Name + NumberInStock + Price + Value).GetHashCode();
            }
        }

        private class ProductValue
        {
            public decimal Value { get; set; }
        }

        private class ProductValueString
        {
            public decimal Value { get; set; }
            public string ValueDefaultAsFormula { get; set; }
            [FormulaResult]
            public string ValueAsString { get; set; }
        }

        [Test]
        public void FetchTest()
        {
            var products = new ExcelMapper(@"..\..\..\products.xlsx").Fetch<Product>().ToList();
            CollectionAssert.AreEqual(new List<Product>
            {
                new Product { Name = "Nudossi", NumberInStock = 60, Price = 1.99m, Value = "C2*D2" },
                new Product { Name = "Halloren", NumberInStock = 33, Price = 2.99m, Value = "C3*D3" },
                new Product { Name = "Filinchen", NumberInStock = 100, Price = 0.99m, Value = "C5*D5" },
            }, products);
        }

        [Test]
        public void FetchWithTypeTest()
        {
            var products = new ExcelMapper(@"..\..\..\products.xlsx").Fetch(typeof(Product));
            CollectionAssert.AreEqual(new List<Product>
            {
                new Product { Name = "Nudossi", NumberInStock = 60, Price = 1.99m, Value = "C2*D2" },
                new Product { Name = "Halloren", NumberInStock = 33, Price = 2.99m, Value = "C3*D3" },
                new Product { Name = "Filinchen", NumberInStock = 100, Price = 0.99m, Value = "C5*D5" },
            }, products);
        }

        [Test]
        public void FetchWithStreamAndIndexTest()
        {
            var stream = new FileStream(@"..\..\..\products.xlsx", FileMode.Open, FileAccess.Read);
            var products = new ExcelMapper().Fetch<Product>(stream, 0);

            CollectionAssert.AreEqual(new List<Product>
            {
                new Product { Name = "Nudossi", NumberInStock = 60, Price = 1.99m, Value = "C2*D2" },
                new Product { Name = "Halloren", NumberInStock = 33, Price = 2.99m, Value = "C3*D3" },
                new Product { Name = "Filinchen", NumberInStock = 100, Price = 0.99m, Value = "C5*D5" },
            }, products);
        }

        [Test]
        public void FetchWithTypeUsingStreamAndIndexTest()
        {
            var stream = new FileStream(@"..\..\..\products.xlsx", FileMode.Open, FileAccess.Read);
            var products = new ExcelMapper().Fetch(stream, typeof(Product), 0);

            CollectionAssert.AreEqual(new List<Product>
            {
                new Product { Name = "Nudossi", NumberInStock = 60, Price = 1.99m, Value = "C2*D2" },
                new Product { Name = "Halloren", NumberInStock = 33, Price = 2.99m, Value = "C3*D3" },
                new Product { Name = "Filinchen", NumberInStock = 100, Price = 0.99m, Value = "C5*D5" },
            }, products);
        }

        [Test]
        public void FetchWithStreamAndSheetNameTest()
        {
            var stream = new FileStream(@"..\..\..\products.xlsx", FileMode.Open, FileAccess.Read);
            var products = new ExcelMapper().Fetch<Product>(stream, "Tabelle1");

            CollectionAssert.AreEqual(new List<Product>
            {
                new Product { Name = "Nudossi", NumberInStock = 60, Price = 1.99m, Value = "C2*D2" },
                new Product { Name = "Halloren", NumberInStock = 33, Price = 2.99m, Value = "C3*D3" },
                new Product { Name = "Filinchen", NumberInStock = 100, Price = 0.99m, Value = "C5*D5" },
            }, products);
        }

        [Test]
        public void FetchWithTypeUsingStreamAndSheetNameTest()
        {
            var stream = new FileStream(@"..\..\..\products.xlsx", FileMode.Open, FileAccess.Read);
            var products = new ExcelMapper().Fetch(stream, typeof(Product), "Tabelle1");

            CollectionAssert.AreEqual(new List<Product>
            {
                new Product { Name = "Nudossi", NumberInStock = 60, Price = 1.99m, Value = "C2*D2" },
                new Product { Name = "Halloren", NumberInStock = 33, Price = 2.99m, Value = "C3*D3" },
                new Product { Name = "Filinchen", NumberInStock = 100, Price = 0.99m, Value = "C5*D5" },
            }, products);
        }

        [Test]
        public void FetchWithFileAndSheetNameTest()
        {
            var products = new ExcelMapper().Fetch<Product>(@"..\..\..\products.xlsx", "Tabelle1");

            CollectionAssert.AreEqual(new List<Product>
            {
                new Product { Name = "Nudossi", NumberInStock = 60, Price = 1.99m, Value = "C2*D2" },
                new Product { Name = "Halloren", NumberInStock = 33, Price = 2.99m, Value = "C3*D3" },
                new Product { Name = "Filinchen", NumberInStock = 100, Price = 0.99m, Value = "C5*D5" },
            }, products);
        }

        [Test]
        public void FetchWithFileAndIndexTest()
        {
            var products = new ExcelMapper().Fetch<Product>(@"..\..\..\products.xlsx", 0);

            CollectionAssert.AreEqual(new List<Product>
            {
                new Product { Name = "Nudossi", NumberInStock = 60, Price = 1.99m, Value = "C2*D2" },
                new Product { Name = "Halloren", NumberInStock = 33, Price = 2.99m, Value = "C3*D3" },
                new Product { Name = "Filinchen", NumberInStock = 100, Price = 0.99m, Value = "C5*D5" },
            }, products);
        }

        [Test]
        public void FetchWithTypeThrowsExceptionWithPrimitivesTest()
        {
            var excel = new ExcelMapper(@"..\..\..\products.xlsx");
            Assert.Throws<ArgumentException>(() => excel.Fetch(typeof(string)));
            Assert.Throws<ArgumentException>(() => excel.Fetch(typeof(object)));
            Assert.Throws<ArgumentException>(() => excel.Fetch(typeof(int)));
            Assert.Throws<ArgumentException>(() => excel.Fetch(typeof(double?)));
        }

        [Test]
        public void FetchValueTest()
        {
            var products = new ExcelMapper(@"..\..\..\products.xlsx").Fetch<ProductValue>().ToList();
            CollectionAssert.AreEqual(new List<decimal> { 119.4m, 98.67m, 99m }, products.Select(p => p.Value).ToList());
        }

        [Test]
        public void FetchValueWithTypeTest()
        {
            var products = new ExcelMapper(@"..\..\..\products.xlsx").Fetch(typeof(ProductValue))
                                                                     .OfType<ProductValue>()
                                                                     .ToList();
            CollectionAssert.AreEqual(new List<decimal> { 119.4m, 98.67m, 99m }, products.Select(p => p.Value).ToList());
        }

        private class ProductException : Product
        {
            public bool Offer { get; set; }
        }

        [Test]
        public void FetchExcpetionWhenEmptyTest()
        {
            var ex = Assert.Throws<ExcelMapperConvertException>(() => new ExcelMapper(@"..\..\..\productsExceptionEmpty.xlsx").Fetch<ProductException>().ToList());
            Assert.That(ex.Message.Contains("<EMPTY>"));
            Assert.That(ex.Message.Contains("[L:1]:[C:2]"));
        }

        [Test]
        public void FetchWithTypeExcpetionWhenEmptyTest()
        {
            var ex = Assert.Throws<ExcelMapperConvertException>(() => new ExcelMapper(@"..\..\..\productsExceptionEmpty.xlsx").Fetch(typeof(ProductException))
                                                                                                                              .OfType<ProductException>()
                                                                                                                              .ToList());
            Assert.That(ex.Message.Contains("<EMPTY>"));
            Assert.That(ex.Message.Contains("[L:1]:[C:2]"));
        }

        [Test]
        public void FetchExcpetionWhenFieldTooBigTest()
        {
            var ex = Assert.Throws<ExcelMapperConvertException>(() => new ExcelMapper(@"..\..\..\productsExceptionTooBig.xlsx").Fetch<ProductException>().ToList());
            //2147483649 is Int.MaxValue + 1
            Assert.That(ex.Message.Contains("2147483649"));
            Assert.That(ex.Message.Contains("[L:1]:[C:1]"));
        }

        [Test]
        public void FetchExcpetionWhenFieldInvalidTest()
        {
            var ex = Assert.Throws<ExcelMapperConvertException>(() => new ExcelMapper(@"..\..\..\productsExceptionInvalid.xlsx").Fetch<ProductException>().ToList());
            Assert.That(ex.Message.Contains("FALSEd"));
            Assert.That(ex.Message.Contains("[L:1]:[C:3]"));
        }

        [Test]
        public void FetchExceptionWhenSheetDoesNotExists()
        {
            var ex = Assert.Throws<ArgumentOutOfRangeException>(() => new ExcelMapper(@"..\..\..\productsExceptionInvalid.xlsx").Fetch<ProductException>("this sheet does not exist").ToList());
            Assert.That(ex.Message.Contains("Sheet not found"));
        }

        [Test]
        public void FetchWithTypeThrowsExceptionWhenSheetDoesNotExists()
        {
            var ex = Assert.Throws<ArgumentOutOfRangeException>(() => new ExcelMapper(@"..\..\..\productsExceptionInvalid.xlsx").Fetch(typeof(ProductException), "This is not a exist")
                                                                                                                                .OfType<ProductException>()
                                                                                                                                .ToList());
            Assert.That(ex.Message.Contains("Sheet not found"));
        }

        private class ProductNoHeader
        {
            [Column(1)]
            public string Name { get; set; }
            [Column(3)]
            public int NumberInStock { get; set; }
            [Column(4)]
            public decimal Price { get; set; }

            public override bool Equals(object obj)
            {
                if (!(obj is ProductNoHeader o)) return false;
                return o.Name == Name && o.NumberInStock == NumberInStock && o.Price == Price;
            }

            public override int GetHashCode()
            {
                return (Name + NumberInStock + Price).GetHashCode();
            }
        }

        [Test]
        public void FetchNoHeaderTest()
        {
            var products = new ExcelMapper(@"..\..\..\productsnoheader.xlsx") { HeaderRow = false }.Fetch<ProductNoHeader>("Products").ToList();
            CollectionAssert.AreEqual(new List<ProductNoHeader>
            {
                new ProductNoHeader { Name = "Nudossi", NumberInStock = 60, Price = 1.99m },
                new ProductNoHeader { Name = "Halloren", NumberInStock = 33, Price = 2.99m },
                new ProductNoHeader { Name = "Filinchen", NumberInStock = 100, Price = 0.99m },
            }, products);
        }

        [Test]
        public void FetchWithTypeNoHeaderTest()
        {
            var products = new ExcelMapper(@"..\..\..\productsnoheader.xlsx") { HeaderRow = false }.Fetch(typeof(ProductNoHeader), "Products")
                                                                                                   .OfType<ProductNoHeader>()
                                                                                                   .ToList();
            CollectionAssert.AreEqual(new List<ProductNoHeader>
            {
                new ProductNoHeader { Name = "Nudossi", NumberInStock = 60, Price = 1.99m },
                new ProductNoHeader { Name = "Halloren", NumberInStock = 33, Price = 2.99m },
                new ProductNoHeader { Name = "Filinchen", NumberInStock = 100, Price = 0.99m },
            }, products);
        }

        [Test]
        public void SaveTest()
        {
            var products = new List<Product>
            {
                new Product { Name = "Nudossi", NumberInStock = 60, Price = 1.99m, Value = "C2*D2" },
                new Product { Name = "Halloren", NumberInStock = 33, Price = 2.99m, Value = "C3*D3" },
                new Product { Name = "Filinchen", NumberInStock = 100, Price = 0.99m, Value = "C4*D4" },
            };

            var file = "productssave.xlsx";

            new ExcelMapper().Save(file, products, "Products");

            var productsFetched = new ExcelMapper(file).Fetch<Product>().ToList();

            CollectionAssert.AreEqual(products, productsFetched);
        }

        [Test]
        public void SaveNoHeaderSaveTest()
        {
            var products = new List<ProductNoHeader>
            {
                new ProductNoHeader { Name = "Nudossi", NumberInStock = 60, Price = 1.99m },
                new ProductNoHeader { Name = "Halloren", NumberInStock = 33, Price = 2.99m },
                new ProductNoHeader { Name = "Filinchen", NumberInStock = 100, Price = 0.99m },
            };

            var file = "productsnoheadersave.xlsx";

            new ExcelMapper() { HeaderRow = false }.Save(file, products, "Products");

            var productsFetched = new ExcelMapper(file) { HeaderRow = false }.Fetch<ProductNoHeader>().ToList();

            CollectionAssert.AreEqual(products, productsFetched);
        }

        [Test]
        public void SaveFetchedTest()
        {
            var excel = new ExcelMapper(@"..\..\..\products.xlsx");
            var products = excel.Fetch<Product>().ToList();

            products[2].Price += 1.0m;

            var file = @"productssavefetched.xlsx";

            excel.Save(file, products);

            var productsFetched = new ExcelMapper(file).Fetch<Product>().ToList();

            CollectionAssert.AreEqual(products, productsFetched);
        }

        private class ProductMapped
        {
            public string NameX { get; set; }
            public int NumberX { get; set; }
            public decimal PriceX { get; set; }

            public override bool Equals(object obj)
            {
                if (!(obj is ProductMapped o)) return false;
                return o.NameX == NameX && o.NumberX == NumberX && o.PriceX == PriceX;
            }

            public override int GetHashCode()
            {
                return (NameX + NumberX + PriceX).GetHashCode();
            }
        }

        [Test]
        public void SaveTrackedObjectsTest()
        {
            var excel = new ExcelMapper(@"..\..\..\products.xlsx") { TrackObjects = true };

            excel.AddMapping(typeof(ProductMapped), "Name", "NameX");
            excel.AddMapping<ProductMapped>("Number", p => p.NumberX);
            excel.AddMapping<ProductMapped>("Price", p => p.PriceX);

            var products = excel.Fetch<ProductMapped>().ToList();

            products[1].PriceX += 1.0m;

            var file = @"productssavetracked.xlsx";

            excel.Save(file);

            var productsFetched = new ExcelMapper(file).Fetch<ProductMapped>().ToList();

            CollectionAssert.AreEqual(products, productsFetched);
        }

        private class GetterSetterProduct
        {
            public string Name { get; set; }
            public DateTime? OfferEnd { get; set; }

            public override bool Equals(object obj)
            {
                if (!(obj is GetterSetterProduct o)) return false;
                return o.Name == Name && o.OfferEnd == OfferEnd;
            }

            public override int GetHashCode()
            {
                return (Name + OfferEnd).GetHashCode();
            }
        }

        [Test]
        public void GetterSetterTest()
        {
            var excel = new ExcelMapper(@"..\..\..\productsconvert.xlsx") { TrackObjects = true };

            excel.AddMapping<GetterSetterProduct>("Name", p => p.Name);
            excel.AddMapping<GetterSetterProduct>("OfferEnd", p => p.OfferEnd)
                .SetCellUsing((c, o) =>
                {
                    if (o == null) c.SetCellValue("NULL"); else c.SetCellValue((DateTime)o);
                })
                .SetPropertyUsing(v =>
                {
                    if ((v as string) == "NULL") return null;
                    return Convert.ChangeType(v, typeof(DateTime), CultureInfo.InvariantCulture);
                });

            var products = excel.Fetch<GetterSetterProduct>().ToList();

            var file = @"productsconverttracked.xlsx";

            excel.Save(file);

            var productsFetched = new ExcelMapper(file).Fetch<GetterSetterProduct>().ToList();

            CollectionAssert.AreEqual(products, productsFetched);
        }

        private class IgnoreProduct
        {
            public string Name { get; set; }
            [Ignore]
            public int Number { get; set; }
            public decimal Price { get; set; }
            public bool Offer { get; set; }
            public DateTime OfferEnd { get; set; }

            public override bool Equals(object obj)
            {
                if (!(obj is IgnoreProduct o)) return false;
                return o.Name == Name && o.Number == Number && o.Price == Price && o.Offer == Offer && o.OfferEnd == OfferEnd;
            }

            public override int GetHashCode()
            {
                return (Name + Number + Price + Offer + OfferEnd).GetHashCode();
            }
        }

        [Test]
        public void IgnoreTest()
        {
            var excel = new ExcelMapper(@"..\..\..\products.xlsx");
            excel.Ignore<IgnoreProduct>(p => p.Price);
            var products = excel.Fetch<IgnoreProduct>().ToList();

            var nudossi = products[0];
            Assert.AreEqual("Nudossi", nudossi.Name);
            Assert.AreEqual(0, nudossi.Number);
            Assert.AreEqual(0m, nudossi.Price);
            Assert.IsFalse(nudossi.Offer);

            var halloren = products[1];
            Assert.IsTrue(halloren.Offer);
            Assert.AreEqual(new DateTime(2015, 12, 31), halloren.OfferEnd);

            var file = "productsignored.xlsx";

            new ExcelMapper().Save(file, products, "Products");

            var productsFetched = new ExcelMapper(file).Fetch<IgnoreProduct>().ToList();

            CollectionAssert.AreEqual(products, productsFetched);
        }

        private class NullableProduct
        {
            public string Name { get; set; }
            public int? Number { get; set; }
            public decimal? Price { get; set; }
            public bool? Offer { get; set; }
            public DateTime? OfferEnd { get; set; }

            public override bool Equals(object obj)
            {
                if (!(obj is NullableProduct o)) return false;
                return o.Name == Name && o.Number == Number && o.Price == Price && o.Offer == Offer && o.OfferEnd == OfferEnd;
            }

            public override int GetHashCode()
            {
                return (Name + Number + Price + Offer + OfferEnd).GetHashCode();
            }
        }

        [Test]
        public void NullableTest()
        {
            var workbook = WorkbookFactory.Create(@"..\..\..\products.xlsx");
            var excel = new ExcelMapper(workbook);
            var products = excel.Fetch<NullableProduct>().ToList();

            var nudossi = products[0];
            Assert.AreEqual("Nudossi", nudossi.Name);
            Assert.AreEqual(60, nudossi.Number);
            Assert.AreEqual(1.99m, nudossi.Price);
            Assert.IsFalse(nudossi.Offer.Value);
            nudossi.OfferEnd = null;

            var halloren = products[1];
            Assert.IsTrue(halloren.Offer.Value);
            Assert.AreEqual(new DateTime(2015, 12, 31), halloren.OfferEnd);
            halloren.Number = null;
            halloren.Offer = null;

            var file = "productsnullable.xlsx";

            new ExcelMapper().Save(file, products, "Products");

            var productsFetched = new ExcelMapper(file).Fetch<NullableProduct>().ToList();

            CollectionAssert.AreEqual(products, productsFetched);
        }

        private class DataFormatProduct
        {
            [DataFormat(0xf)]
            public DateTime Date { get; set; }

            [DataFormat("0%")]
            public decimal Number { get; set; }
        }

        [Test]
        public void DataFormatTest()
        {
            var p = new DataFormatProduct { Date = new DateTime(2015, 12, 31), Number = 0.47m };
            var file = "productsdataformat.xlsx";
            new ExcelMapper().Save(file, new[] { p });
            var pfs = new ExcelMapper(file).Fetch<DataFormatProduct>().ToList();

            Assert.AreEqual(1, pfs.Count);
            var pf = pfs[0];
            Assert.AreEqual(p.Date, pf.Date);
            Assert.AreEqual(p.Number, pf.Number);
        }

        private class DataItem
        {
            [Ignore]
            public string Bql { get; set; }

            [Ignore]
            public int Id { get; set; }

            [Column(2)]
            public string OriginalBql { get; set; }

            [Column(1)]
            public string Title { get; set; }

            [Column(3)]
            public string TranslatedBql { get; set; }

            public override bool Equals(object obj)
            {
                if (!(obj is DataItem o)) return false;
                return o.Bql == Bql && o.Id == Id && o.OriginalBql == OriginalBql && o.Title == Title && o.TranslatedBql == TranslatedBql;
            }

            public override int GetHashCode()
            {
                return (Bql + Id + OriginalBql + Title + TranslatedBql).GetHashCode();
            }
        }

        [Test]
        public void ColumnTest()
        {
            var excel = new ExcelMapper(@"..\..\..\dataitems.xlsx") { HeaderRow = false };
            var items = excel.Fetch<DataItem>().ToList();

            var trackedFile = "dataitemstracked.xlsx";
            excel.Save(trackedFile, "DataItems");
            var itemsTracked = excel.Fetch<DataItem>(trackedFile, "DataItems").ToList();
            CollectionAssert.AreEqual(items, itemsTracked);

            var saveFile = "dataitemssave.xlsx";
            new ExcelMapper().Save(saveFile, items, "DataItems");
            var itemsSaved = new ExcelMapper().Fetch<DataItem>(saveFile, "DataItems").ToList();
            CollectionAssert.AreEqual(items, itemsSaved);
        }

        [Test]
        public void ColumnTestUsingFetchWithType()
        {
            var excel = new ExcelMapper(@"..\..\..\dataitems.xlsx") { HeaderRow = false };
            var items = excel.Fetch(typeof(DataItem)).OfType<DataItem>().ToList();

            var trackedFile = "dataitemstracked.xlsx";
            excel.Save(trackedFile, "DataItems");
            var itemsTracked = excel.Fetch(trackedFile, typeof(DataItem), "DataItems").OfType<DataItem>().ToList();
            CollectionAssert.AreEqual(items, itemsTracked);

            var saveFile = "dataitemssave.xlsx";
            new ExcelMapper().Save(saveFile, items, "DataItems");
            var itemsSaved = new ExcelMapper().Fetch(saveFile, typeof(DataItem), "DataItems").OfType<DataItem>().ToList();
            CollectionAssert.AreEqual(items, itemsSaved);
        }

        [Test]
        public void FetchMinMaxTest()
        {
            var products = new ExcelMapper(@"..\..\..\ProductsMinMaxRow.xlsx")
            {
                HeaderRowNumber = 2,
                MinRowNumber = 6,
                MaxRowNumber = 9,
            }.Fetch<Product>().ToList();
            CollectionAssert.AreEqual(new List<Product>
            {
                new Product { Name = "Nudossi", NumberInStock = 60, Price = 1.99m, Value = "C7*D7" },
                new Product { Name = "Halloren", NumberInStock = 33, Price = 2.99m, Value = "C8*D8" },
                new Product { Name = "Filinchen", NumberInStock = 100, Price = 0.99m, Value = "C10*D10" },
            }, products);
        }

        [Test]
        public void SaveMinMaxTest()
        {
            var products = new List<Product>
            {
                new Product { Name = "Nudossi", NumberInStock = 60, Price = 1.99m, Value = "C2*D2" },
                new Product { Name = "Halloren", NumberInStock = 33, Price = 2.99m, Value = "C3*D3" },
                new Product { Name = "Filinchen", NumberInStock = 100, Price = 0.99m, Value = "C4*D4" },
            };

            var file = "productsminmaxsave.xlsx";

            new ExcelMapper
            {
                HeaderRowNumber = 1,
                MinRowNumber = 3
            }.Save(file, products, "Products");

            var productsFetched = new ExcelMapper(file)
            {
                HeaderRowNumber = 1,
                MinRowNumber = 3
            }.Fetch<Product>().ToList();

            CollectionAssert.AreEqual(products, productsFetched);
        }

        [Test]
        public void FormulaResultAttributeTest()
        {
            var products = new ExcelMapper(@"..\..\..\ProductsAsString.xlsx").Fetch<ProductValueString>().ToList();
            CollectionAssert.AreEqual(new List<string> { "119.4", "98.67", "99" }, products.Select(p => p.ValueAsString).ToList());
        }

        private class ProductFormulaMapped
        {
            public decimal Result { get; set; }
            public string Formula { get; set; }
            public string ResultString { get; set; }

            public override bool Equals(object obj)
            {
                if (!(obj is ProductFormulaMapped o)) return false;
                return o.Result == Result && o.Formula == Formula && o.ResultString == ResultString;
            }

            public override int GetHashCode()
            {
                return (Formula + ResultString + Result).GetHashCode();
            }
        }

        [Test]
        public void FormulaResultMappedTest()
        {
            var excel = new ExcelMapper(@"..\..\..\ProductsAsString.xlsx");

            excel.AddMapping<ProductFormulaMapped>("Value", p => p.Result);
            excel.AddMapping<ProductFormulaMapped>("ValueDefaultAsFormula", p => p.Formula);
            excel.AddMapping<ProductFormulaMapped>("ValueAsString", p => p.ResultString).AsFormulaResult();

            var products = excel.Fetch<ProductFormulaMapped>().ToList();
            var expectedProducts = new List<ProductFormulaMapped>
            {
                new ProductFormulaMapped { Result = 119.4m, Formula = "C2*D2", ResultString = "119.4" },
                new ProductFormulaMapped { Result = 98.67m, Formula = "C3*D3", ResultString = "98.67" },
                new ProductFormulaMapped { Result = 99m, Formula = "C5*D5", ResultString = "99" },
            };

            Assert.AreEqual(expectedProducts[0], products[0]);

            CollectionAssert.AreEqual(expectedProducts, products);
        }

        [Test]
        public void TestExcelMapperConvertException()
        {
            ExcelMapperConvertException ex =
                new ExcelMapperConvertException("cellvalue", typeof(string), 12, 34);

            // Sanity check: Make sure custom properties are set before serialization
            Assert.AreEqual(12, ex.Line);
            Assert.AreEqual(34, ex.Column);

            // Round-trip the exception: Serialize and de-serialize with a BinaryFormatter
            var bf = new BinaryFormatter();
            using (var ms = new MemoryStream())
            {
                // "Save" object state
                bf.Serialize(ms, ex);

                // Re-use the same stream for de-serialization
                ms.Seek(0, 0);

                // Replace the original exception with de-serialized one
                ex = (ExcelMapperConvertException)bf.Deserialize(ms);
            }

            // Make sure custom properties are preserved after serialization
            Assert.AreEqual(12, ex.Line);
            Assert.AreEqual(34, ex.Column);

            Assert.Throws<ArgumentNullException>(() => ex.GetObjectData(null, new System.Runtime.Serialization.StreamingContext()));
        }
    }
}
