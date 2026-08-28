using OfficeCli.Core;
using System.Text.Json;
using Xunit;

namespace OfficeCli.Tests;

public class DelimitedTextTests
{
    [Fact]
    public void TypedJsonGridPreservesFormattedThousandsAsOneCell()
    {
        var rows = DelimitedText.ParseJsonGrid(
            "[[\"指标\",\"结果\"],[\"总会话量\",{\"value\":12480,\"display\":\"12,480\"}]]");

        Assert.Equal(2, rows.Length);
        Assert.Equal(2, rows[1].Length);
        Assert.Equal("12,480", rows[1][1]);
    }

    [Fact]
    public void RaggedDelimitedRowsAreRejectedInsteadOfPadded()
    {
        var rows = DelimitedText.ParseGrid("指标,结果,说明;总会话量,12,480,会话规模", ',', ';');

        var error = Assert.Throws<ArgumentException>(() => DelimitedText.EnsureRectangular(rows));
        Assert.Contains("row 2 has 4 cell(s)", error.Message);
        Assert.Contains("dataJson", error.Message);
    }

    [Fact]
    public void TypedJsonGridRejectsRaggedRows()
    {
        var error = Assert.Throws<ArgumentException>(() =>
            DelimitedText.ParseJsonGrid("[[\"A\",\"B\"],[1,2,3]]"));

        Assert.Contains("ragged rows are not padded", error.Message);
    }

    [Fact]
    public void BatchPropsPreserveNaturalDataJsonArraysWithoutSplittingThousands()
    {
        const string payload = """
            [{
              "command": "add",
              "parent": "/body",
              "type": "table",
              "props": {
                "dataJson": [
                  ["指标", "结果"],
                  ["总会话量", {"value": 12480, "display": "12,480"}]
                ]
              }
            }]
            """;

        var items = JsonSerializer.Deserialize<List<BatchItem>>(payload);

        var dataJson = Assert.Single(items!).Props!["dataJson"];
        var rows = DelimitedText.ParseJsonGrid(dataJson);
        Assert.Equal("12,480", rows[1][1]);
        Assert.Equal(2, rows[1].Length);
    }
}
