using System.Net.Http;
using System.Threading.Tasks;
using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// The launch check-in's form, which both ends share (2026-10-03). The server's side - reading this
// very form - is CRT.Server.Tests' CheckInFormReaderTests; this pins the form's shape.
// ###########################################################################################
public sealed class CheckInContractTests
{
    // *** THE PHP PAGE'S FORM, BYTE FOR BYTE. *** Every CRT already installed posts exactly this to
    // /app-checkin/, which Apache forwards to the server - so a field renamed here would still
    // compile on both ends, and every older CRT's check-in would stop counting.
    [Fact]
    public async Task The_form_is_the_PHP_pages_four_fields_url_encoded()
    {
        using FormUrlEncodedContent form = CheckInContract.BuildForm("Windows", "Microsoft Windows 10.0.19045", "64-bit");

        Assert.Equal("application/x-www-form-urlencoded", form.Headers.ContentType!.MediaType);
        Assert.Equal(
            "control=CRT&osHighlevel=Windows&osVersion=Microsoft+Windows+10.0.19045&cpu=64-bit",
            await form.ReadAsStringAsync());
    }

    [Fact]
    public async Task A_missing_value_is_sent_empty_rather_than_left_out()
    {
        using FormUrlEncodedContent form = CheckInContract.BuildForm(null!, null!, null!);

        Assert.Equal("control=CRT&osHighlevel=&osVersion=&cpu=", await form.ReadAsStringAsync());
    }

    [Fact]
    public void The_route_sits_under_usage_beside_board_views()
    {
        Assert.Equal("usage/check-in", CheckInContract.PathUnderApi);
    }
}
