using System;
using System.Collections.Generic;

namespace MultipartExcel
{
    public static class DataProvider
    {
        public const int ChunkSize = 500000;

        public static IEnumerable<TestData> CreateTestDataChunk()
        {
            for (var i = 0; i < ChunkSize; i++)
            {
                yield return new TestData
                {
                    Id = Guid.NewGuid(),
                    Name = Convert.ToBase64String(Guid.NewGuid().ToByteArray()),
                    CreateDate = DateTime.UtcNow.AddHours(-2),
                    UpdateDate = i % 2 == 0 ? (DateTime?) null : DateTime.UtcNow,
                    Count = i
                };
            }
        }
    }

    public class TestData
    {
        public Guid Id { get; set; }
        public string Name { get; set; }
        public DateTime CreateDate { get; set; }
        public DateTime? UpdateDate { get; set; }
        public int Count { get; set; }
    }
}