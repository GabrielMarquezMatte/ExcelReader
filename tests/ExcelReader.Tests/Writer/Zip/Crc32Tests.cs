using System.Diagnostics.CodeAnalysis;
using System.Runtime.Intrinsics.X86;
using ExcelReader.Core.Writer.Zip;
using ArmCrc32 = System.Runtime.Intrinsics.Arm.Crc32;

namespace ExcelReader.Tests.Writer.Zip
{
    public class Crc32Tests
    {
        private delegate uint AppendFunction(uint crc, ReadOnlySpan<byte> data);

        internal static uint BitwiseCrc(ReadOnlySpan<byte> data)
        {
            uint register = uint.MaxValue;
            foreach (byte value in data)
            {
                register ^= value;
                for (int bit = 0; bit < 8; bit++)
                {
                    uint mask = 0u - (register & 1);
                    register = (register >> 1) ^ (0xEDB8_8320 & mask);
                }
            }
            return ~register;
        }

        private static List<(string Name, AppendFunction Append)> AvailablePaths()
        {
            var paths = new List<(string Name, AppendFunction Append)>
            {
                ("dispatch", Crc32.Append),
                ("table", Crc32.AppendTable),
            };
            if (Pclmulqdq.IsSupported)
            {
                paths.Add(("fold128", Crc32.AppendFold128));
            }
            if (Pclmulqdq.V256.IsSupported)
            {
                paths.Add(("fold256", Crc32.AppendFold256));
            }
            if (ArmCrc32.Arm64.IsSupported)
            {
                paths.Add(("arm64", Crc32.AppendArm64));
            }
            return paths;
        }

        [Fact]
        public void EveryPathGivesTheStandardCheckValue()
        {
            foreach ((string name, AppendFunction append) in AvailablePaths())
            {
                uint actual = append(0, "123456789"u8);
                Assert.True(actual == 0xCBF4_3926, $"{name}: {actual:X8}");
            }
        }

        [Fact]
        public void EmptyInputLeavesTheCrcUnchanged()
        {
            foreach ((string name, AppendFunction append) in AvailablePaths())
            {
                uint actual = append(0x1234_5678, []);
                Assert.True(actual == 0x1234_5678, $"{name}: {actual:X8}");
            }
        }

        [Fact]
        [SuppressMessage("Security", "CA5394:Do not use insecure randomness",
            Justification = "A seeded generator picks reproducible test inputs; nothing here is secret.")]
        public void EveryPathMatchesTheBitwiseDefinitionWholeAndSplit()
        {
            var random = new Random(5);
            byte[] buffer = new byte[70_064];
            random.NextBytes(buffer);
            List<(string Name, AppendFunction Append)> paths = AvailablePaths();
            for (int trial = 0; trial < 900; trial++)
            {
                int start = random.Next(0, 64);
                int length = trial < 600 ? trial : random.Next(600, 70_000);
                ReadOnlySpan<byte> slice = buffer.AsSpan(start, length);
                uint expected = BitwiseCrc(slice);
                int cut = random.Next(0, length + 1);
                foreach ((string name, AppendFunction append) in paths)
                {
                    uint whole = append(0, slice);
                    uint split = append(append(0, slice[..cut]), slice[cut..]);
                    Assert.True(whole == expected, $"{name} whole, length {length}: {whole:X8} != {expected:X8}");
                    Assert.True(split == expected, $"{name} split at {cut} of {length}: {split:X8} != {expected:X8}");
                }
            }
        }
    }
}
