using ExcelReader.Core.Crypto;
using ExcelReader.Core.Reader;

namespace ExcelReader.Tests
{
    public class EncryptedFixtureInventoryTests
    {
        // Keyed by exact fixture name (not a naming-convention prefix) so an unregistered future
        // fixture fails loudly here instead of the test silently accepting whatever descriptor type
        // it happens to parse to.
        private static readonly Dictionary<string, Type> ExpectedDescriptorTypes = new(StringComparer.Ordinal)
        {
            ["agile-aes256-sha512.xlsx"] = typeof(AgileDescriptor),
            ["agile-aes256-sha512.xlsb"] = typeof(AgileDescriptor),
            ["standard-aes128-sha1.xlsx"] = typeof(StandardDescriptor),
        };

        public static TheoryData<string> Fixtures()
        {
            var data = new TheoryData<string>();
            foreach (string name in EncryptedFixtures.All)
            {
                data.Add(name);
            }
            return data;
        }

        // Guards against a fixture swap silently changing what's actually under test: without this,
        // swapping standard-aes128-sha1.xlsx for an agile file would leave every other encrypted
        // suite green while third-party standard-encryption coverage quietly disappeared.
        [Theory]
        [MemberData(nameof(Fixtures))]
        public void Should_Parse_To_Expected_Descriptor_Type_When_Encrypted(string name)
        {
            using FileStream fs = File.OpenRead(EncryptedFixtures.Path_(name));
            using CfbContainer cfb = CfbContainer.Parse(fs, ownsSource: true, ExcelReaderOptions.Default);
            byte[] info = cfb.ReadStream("EncryptionInfo", 64 * 1024);
            EncryptionDescriptor descriptor = EncryptionDescriptor.Parse(info, ExcelReaderOptions.Default);

            Assert.True(ExpectedDescriptorTypes.TryGetValue(name, out Type? expected),
                $"No expected descriptor type registered for fixture '{name}'.");
            Assert.IsType(expected, descriptor);
        }

        [Theory]
        [MemberData(nameof(Fixtures))]
        public void Should_Have_Encrypted_And_Plain_Fixture_When_Named_In_Inventory(string name)
        {
            Assert.True(File.Exists(EncryptedFixtures.Path_(name)), $"missing fixture {name}");
            Assert.True(File.Exists(EncryptedFixtures.PlainPath(name)), $"missing oracle for {name}");
        }

        // A CFB container, not a ZIP: this is exactly why Excel.Open misroutes these files today.
        [Theory]
        [MemberData(nameof(Fixtures))]
        public void Should_Be_Cfb_Container_When_Encrypted(string name)
        {
            byte[] head = EncryptedFixtures.Bytes(name)[..8];
            Assert.Equal(new byte[] { 0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1 }, head);
        }

        // The oracle is a plain ZIP ("PK\x03\x04") - what our decryptor must reproduce.
        [Theory]
        [MemberData(nameof(Fixtures))]
        public void Should_Be_Zip_When_Plain(string name)
        {
            byte[] head = EncryptedFixtures.PlainBytes(name)[..4];
            Assert.Equal(new byte[] { 0x50, 0x4B, 0x03, 0x04 }, head);
        }

        // A single-segment fixture would pass with completely broken segment indexing; both
        // fixtures happen to exceed one segment already, which is why no separate
        // "multisegment" fixture is needed in this pass.
        [Theory]
        [MemberData(nameof(Fixtures))]
        public void Should_Exceed_One_Segment_When_Fixture_Used_For_Boundary_Tests(string name)
        {
            Assert.True(EncryptedFixtures.PlainBytes(name).Length > 4096);
        }
    }
}
