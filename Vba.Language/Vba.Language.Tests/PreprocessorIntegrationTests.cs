using Vba.Language.Preprocessor;

using Xunit;

namespace Vba.Language.Tests
{
    public class PreprocessorIntegrationTests
    {
        private const string ModuleHeader = "Attribute VB_Name = \"Test\"\r\n";

        [Theory]
        [InlineData("", "")]
        [InlineData("#If True Then\r\n#End If", "")]
        [InlineData("#If True Then\r\ntest\r\n#End If", "test\r\n")]
        public void CanPreprocessModule(string source, string expected)
        {
            var preprocessor = new SourceTextReader(ModuleHeader + source);

            var output = preprocessor.ReadToEnd();

            Assert.Equal(expected, output);
        }
    }
}
