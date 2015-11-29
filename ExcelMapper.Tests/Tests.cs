using NUnit.Framework;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Ganss.Excel.Tests
{
    public class Tests
    {
        public class Product
        {
            public string Name { get; set; }
            [Column("Number")]
            public int NumberInStock { get; set; }
            public decimal Price { get; set; }

            public override bool Equals(object obj)
            {
                var o = obj as Product;
                if (o == null) return false;
                return o.Name == Name && o.NumberInStock == NumberInStock && o.Price == Price;
            }

            public override int GetHashCode()
            {
                return (Name + NumberInStock + Price).GetHashCode();
            }
        }

        [Test]
        public void FetchTest()
        {
            var products = new ExcelMapper(@"..\..\products.xlsx").Fetch<Product>().ToList();
            CollectionAssert.AreEqual(new List<Product>
            {
                new Product { Name = "Nudossi", NumberInStock = 60, Price = 1.99m },
                new Product { Name = "Halloren", NumberInStock = 33, Price = 2.99m },
                new Product { Name = "Filinchen", NumberInStock = 100, Price = 0.99m },
            }, products);
        }

        public class ProductNoHeader
        {
            [Column(1)]
            public string Name { get; set; }
            [Column(3)]
            public int NumberInStock { get; set; }
            [Column(4)]
            public decimal Price { get; set; }

            public override bool Equals(object obj)
            {
                var o = obj as ProductNoHeader;
                if (o == null) return false;
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
            var products = new ExcelMapper(@"..\..\productsnoheader.xlsx") { HeaderRow = false }.Fetch<ProductNoHeader>("Products").ToList();
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
                new Product { Name = "Nudossi", NumberInStock = 60, Price = 1.99m },
                new Product { Name = "Halloren", NumberInStock = 33, Price = 2.99m },
                new Product { Name = "Filinchen", NumberInStock = 100, Price = 0.99m },
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
            var excel = new ExcelMapper(@"..\..\products.xlsx");
            var products = excel.Fetch<Product>().ToList();

            products[2].Price += 1.0m;

            var file = @"productssavefetched.xlsx";

            excel.Save(file, products);

            var productsFetched = new ExcelMapper(file).Fetch<Product>().ToList();

            CollectionAssert.AreEqual(products, productsFetched);
        }

        public class ProductMapped
        {
            public string NameX { get; set; }
            public int NumberX { get; set; }
            public decimal PriceX { get; set; }

            public override bool Equals(object obj)
            {
                var o = obj as ProductMapped;
                if (o == null) return false;
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
            var excel = new ExcelMapper(@"..\..\products.xlsx") { TrackObjects = true };

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

        public class IgnoreProduct
        {
            public string Name { get; set; }
            [Ignore]
            public int Number { get; set; }
            public decimal Price { get; set; }
        }

        [Test]
        public void IgnoreTest()
        {
            var excel = new ExcelMapper(@"..\..\products.xlsx");
            excel.Ignore<IgnoreProduct>(p => p.Price);
            var products = excel.Fetch<IgnoreProduct>().ToList();

            var nudossi = products[0];
            Assert.AreEqual("Nudossi", nudossi.Name);
            Assert.AreEqual(0, nudossi.Number);
            Assert.AreEqual(0m, nudossi.Price);
        }
    }
}
