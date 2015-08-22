using Vba.Language.Preprocessor;

using Xunit;

namespace Vba.Language.Tests.Preprocessor
{
    public class PreprocessorIntegrationTests
    {
        [Theory]
        [InlineData("", "")]
        [InlineData("#If True Then\r\n#End If", "")]
        [InlineData("#If True Then\r\nActive Text\r\n#End If", "Active Text\r\n")]
        [InlineData("#If False Then\r\nDeactivated Text\r\n#End If", "")]
        [InlineData("#If False Then\r\nDeactivated Text\r\n#Else\r\nActive Text\r\n#End If", "Active Text\r\n")]
        public void CanPreprocessModule(string source, string expected)
        {
            var preprocessor = new SourceTextReader(source);

            var output = preprocessor.ReadToEnd();

            Assert.Equal(expected, output);
        }
    }
}
