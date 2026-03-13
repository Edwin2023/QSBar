using Xunit;
using QSBar.Core;
using QSBar.Core.Functions;
using QSBar.Core.Lookup;
using QSBar.Core.Formatting;
using QSBar.Core.Photos;
using System;
using System.Collections.Generic;

namespace QSBar.Tests
{
    public class CoreTests
    {
        [Fact]
        public void TextUtil_Normalize_ShouldTrimAndHandleNull()
        {
            Assert.Equal("test", TextUtil.Normalize("  test  "));
            Assert.Equal(string.Empty, TextUtil.Normalize(null));
            Assert.Equal(string.Empty, TextUtil.Normalize("   "));
        }

        [Fact]
        public void PhotoHelper_Scale_ShouldCalculateCorrectly()
        {
            var (w, h) = PhotoHelper.Scale(100, 200, 50);
            Assert.Equal(50, w);
            Assert.Equal(100, h);

            var (w2, h2) = PhotoHelper.Scale(100, 200, 150);
            Assert.Equal(150, w2);
            Assert.Equal(300, h2);
        }

        [Fact]
        public void DateHelper_FormatDate_ShouldFormatCorrectly()
        {
            // Test DateTime
            var dt = new DateTime(2023, 10, 5);
            Assert.Equal("2023-10-5", DateHelper.FormatDate(dt));

            // Test string
            Assert.Equal("2023-10-5", DateHelper.FormatDate("2023/10/05"));
            Assert.Equal("2023-10-5", DateHelper.FormatDate("2023-10-05"));

            // Test OADate (double)
            double oaDate = 45200; // 2023-10-01
            Assert.Equal("2023-10-1", DateHelper.FormatDate(oaDate));

            // Test null/empty
            Assert.Equal("", DateHelper.FormatDate(null));
            Assert.Equal("", DateHelper.FormatDate(""));
        }

        [Fact]
        public void BQFunctions_PriceWithTax_ShouldCalculateCorrectly()
        {
            Assert.Equal(110.0, BQFunctions.PriceWithTax(100.0, 0.1), 10);
            Assert.Equal(105.0, BQFunctions.PriceWithTax(100.0, 0.05), 10);
        }

        [Fact]
        public void BQFunctions_SumRange_ShouldSumCorrectly()
        {
            object[,] data = new object[,] { { 1.0, 2.0 }, { "3", 4.0 } };
            Assert.Equal(10.0, BQFunctions.SumRange(data));
        }

        [Fact]
        public void BQFunctions_WeightedAverage_ShouldCalculateCorrectly()
        {
            object[,] prices = new object[,] { { 10.0 }, { 20.0 } };
            object[,] quantities = new object[,] { { 1.0 }, { 2.0 } };
            // (10*1 + 20*2) / (1+2) = 50 / 3 = 16.666...
            Assert.Equal(16.666666666666668, BQFunctions.WeightedAverage(prices, quantities));
        }

        [Fact]
        public void BQFunctions_GrossMargin_ShouldCalculateCorrectly()
        {
            Assert.Equal(0.2, BQFunctions.GrossMargin(100.0, 80.0));
            Assert.Equal(0, BQFunctions.GrossMargin(0, 80.0));
        }

        [Fact]
        public void AdvancedLookup_SmartLookupFromArrays_ShouldMatchCorrectly()
        {
            object[,] table = new object[,] { { "Apple Inc." }, { "Microsoft Corp" }, { "Google" } };
            object[,] returns = new object[,] { { "AAPL" }, { "MSFT" }, { "GOOG" } };

            // Exact match
            Assert.Equal("AAPL", AdvancedLookup.SmartLookupFromArrays("Apple Inc.", table, returns, 0.8));

            // Fuzzy match (contain)
            Assert.Equal("AAPL", AdvancedLookup.SmartLookupFromArrays("Apple", table, returns, 0.4));

            // Keyword match
            Assert.Equal("MSFT", AdvancedLookup.SmartLookupFromArrays("Microsoft", table, returns, 0.4));
        }

        [Fact]
        public void AdvancedLookup_ExtractKeywords_ShouldWork()
        {
            string text = "THE Apple 10MM THICK AND Microsoft";
            var keywords = AdvancedLookup.ExtractKeywords(text);
            
            // "THE", "10MM", "THICK", "AND" should be removed
            Assert.Contains("Apple", keywords);
            Assert.Contains("Microsoft", keywords);
            Assert.DoesNotContain("THE", keywords);
            Assert.DoesNotContain("10MM", keywords);
        }

        [Fact]
        public void RowLevelDetector_ComputeGroups_ShouldWork()
        {
            // Levels: 0, 1, 1, 2, 2, 1, 0
            // Groups expected:
            // (1, 5, 1) - because 1 starts at index 1 and drops at index 6
            // (3, 4, 2) - because 2 starts at index 3 and drops at index 5
            int[] levels = { 0, 1, 1, 2, 2, 1, 0 };
            var groups = RowLevelDetector.ComputeGroups(levels);

            Assert.Contains(groups, g => g.Start == 1 && g.End == 5 && g.Level == 1);
            Assert.Contains(groups, g => g.Start == 3 && g.End == 4 && g.Level == 2);
        }
    }
}
