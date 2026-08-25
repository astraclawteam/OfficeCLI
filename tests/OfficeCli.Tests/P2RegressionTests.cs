using OfficeCli.Core;
using Xunit;

namespace OfficeCli.Tests;

public sealed class P2RegressionTests
{
    [Theory]
    [InlineData("1")]
    [InlineData("true")]
    [InlineData("YES")]
    [InlineData(" on ")]
    public void MetadataFingerprintHasExplicitOptOut_Issue337(string value)
        => Assert.True(OfficeCliMetadata.FingerprintDisabled(value));

    [Fact]
    public void UnixConfigUsesXdgWithoutBreakingLegacyDirectory_Issue300()
    {
        using var temp = new TempDirectory();
        var xdg = temp.File("xdg");
        Assert.Equal(Path.Combine(xdg, "officecli"),
            UpdateChecker.ResolveConfigDir(temp.Path, xdg, isWindows: false));

        var legacy = temp.File(".officecli");
        Directory.CreateDirectory(legacy);
        Assert.Equal(legacy,
            UpdateChecker.ResolveConfigDir(temp.Path, xdg, isWindows: false));
    }
}
