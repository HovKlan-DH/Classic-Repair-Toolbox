using Avalonia;
using Handlers.DataHandling;
using Handlers.Geometry;
using System;
using System.Collections.Generic;
using System.Linq;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// Where a part picture's pieces stand (PartPictureGeometry; owner request, 2026-11-09). A part lying across has a lead out
// of each end and its labels under the leads' outer ends; a part standing has its legs at an even pitch under its body
// and each label under its foot. Every piece inside the picture, the marks on the side the polarity says, and a lead
// found by pointing at it.
// ###########################################################################################
public sealed class PartPictureGeometryTests
{
    private static PartPicture.Picture Picture(PartPicture.Shape shape, int leads, int? marked = null, params string[] names) =>
        new(
            shape,
            Enumerable.Range(1, leads).Select(number => new PartPicture.Lead(number, names.Length >= number ? names[number - 1] : string.Empty)).ToList(),
            marked,
            "10µF\n16V",
            string.Empty);

    private static readonly PartPicture.Shape[] Lying =
    [
        PartPicture.Shape.AxialElectrolytic,
        PartPicture.Shape.AxialCapacitor,
        PartPicture.Shape.GlassDiode,
        PartPicture.Shape.EpoxyDiode,
        PartPicture.Shape.FerriteBead,
        PartPicture.Shape.Coil,
    ];

    // Every shape, with the leads it is drawn with - some with two counts.
    private static readonly (PartPicture.Shape Shape, int Leads)[] Cases =
    [
        (PartPicture.Shape.RadialElectrolytic, 2),
        (PartPicture.Shape.AxialElectrolytic, 2),
        (PartPicture.Shape.DiscCapacitor, 2),
        (PartPicture.Shape.RadialCapacitor, 2),
        (PartPicture.Shape.AxialCapacitor, 2),
        (PartPicture.Shape.GlassDiode, 2),
        (PartPicture.Shape.EpoxyDiode, 2),
        (PartPicture.Shape.TO92, 3),
        (PartPicture.Shape.TO220, 3),
        (PartPicture.Shape.Crystal, 2),
        (PartPicture.Shape.Crystal, 3),
        (PartPicture.Shape.FerriteBead, 2),
        (PartPicture.Shape.Coil, 2),
        (PartPicture.Shape.EmiFilter, 3),
        (PartPicture.Shape.ResistorNetwork, 10),
        (PartPicture.Shape.PinHeader, 1),
        (PartPicture.Shape.PinHeader, 7),
    ];

    public static TheoryData<PartPicture.Shape, int> Every_shape()
    {
        var data = new TheoryData<PartPicture.Shape, int>();

        foreach ((PartPicture.Shape shape, int leads) in Cases)
        {
            data.Add(shape, leads);
        }

        return data;
    }

    // A shape added to PartPicture is checked here too.
    [Fact]
    public void Every_shape_has_a_test_case_here()
    {
        Assert.Equal(Enum.GetValues<PartPicture.Shape>().Order(), Cases.Select(row => row.Shape).Distinct().Order());
    }

    // Nothing drawn or written outside the picture's size - the body, its pieces, every wire and label.
    [Theory]
    [MemberData(nameof(Every_shape))]
    public void Every_piece_stands_inside_the_picture(PartPicture.Shape shape, int leads)
    {
        PartPictureGeometry.Layout layout = PartPictureGeometry.For(Picture(shape, leads, marked: leads == 2 ? 1 : null));
        var picture = new Rect(layout.Size);

        Assert.Equal(leads, layout.Leads.Count);
        Assert.True(picture.Contains(layout.Body), $"{shape}: body {layout.Body} outside {picture}");

        foreach (Rect piece in new[] { layout.Band, layout.Tab, layout.Hole }.OfType<Rect>().Append(layout.Marking))
        {
            Assert.True(picture.Contains(piece), $"{shape}: {piece} outside {picture}");
        }

        foreach (PartPictureGeometry.LeadPlace lead in layout.Leads)
        {
            Assert.True(picture.Contains(lead.Wire), $"{shape}: wire {lead.Wire} outside {picture}");
            Assert.True(picture.Contains(lead.Label), $"{shape}: label {lead.Label} outside {picture}");
            Assert.True(picture.Contains(lead.NameBox), $"{shape}: name {lead.NameBox} outside {picture}");
        }
    }

    // A part lying across: a lead out of each end of the body, from the picture's edge, and the labels under them.
    [Fact]
    public void A_part_lying_across_has_a_lead_out_of_each_end_and_its_labels_under_them()
    {
        foreach (PartPicture.Shape shape in Lying)
        {
            PartPictureGeometry.Layout layout = PartPictureGeometry.For(Picture(shape, 2, marked: 0));
            (PartPictureGeometry.LeadPlace left, PartPictureGeometry.LeadPlace right) = (layout.Leads[0], layout.Leads[1]);

            Assert.True(layout.IsAxial);
            // Its leads only as long as their labels need (owner request, 2026-11-09: "there is no reason to have such long
            // legs, as it does take space from the right-side table/info").
            Assert.Equal(layout.Body.Width + (2 * (PartPictureGeometry.Margin + PartPictureGeometry.AxialLeadOut)), layout.Size.Width, 3);
            Assert.True(left.Wire.Width <= PartPictureGeometry.Margin + PartPictureGeometry.AxialLeadOut + 2, $"{shape}: lead {left.Wire.Width}");
            Assert.True(left.Wire.Left < layout.Body.Left && left.Wire.Right >= layout.Body.Left, $"{shape}: left lead {left.Wire}");
            Assert.True(right.Wire.Right > layout.Body.Right && right.Wire.Left <= layout.Body.Right, $"{shape}: right lead {right.Wire}");
            Assert.Equal(layout.Body.Center.Y, left.Wire.Center.Y, 3);

            // Under its lead, clear of the body, the name under the number.
            Assert.True(left.Label.Top >= left.Wire.Bottom && left.Label.Right <= layout.Body.Left, shape.ToString());
            Assert.True(right.Label.Top >= right.Wire.Bottom && right.Label.Left >= layout.Body.Right, shape.ToString());
            Assert.True(left.NameBox.Top >= left.Label.Bottom);

            // Each lead's area is the picture beyond its end of the body.
            Assert.Equal(new Rect(0, 0, layout.Body.Left, layout.Size.Height), left.Area);
            Assert.Equal(layout.Body.Right, right.Area.Left);
            Assert.Equal(layout.Size.Width, right.Area.Right, 3);
        }
    }

    // A diode's band at the cathode's end.
    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    public void A_diodes_band_stands_at_its_cathodes_end(int marked)
    {
        foreach (PartPicture.Shape shape in new[] { PartPicture.Shape.GlassDiode, PartPicture.Shape.EpoxyDiode })
        {
            PartPictureGeometry.Layout layout = PartPictureGeometry.For(Picture(shape, 2, marked));
            Rect band = layout.Band!.Value;

            Assert.Equal(marked == 0, band.Center.X < layout.Body.Center.X);
            Assert.True(layout.Body.Contains(band));
        }
    }

    // ###########################################################################################
    // An axial electrolytic has no minus band (owner request, 2026-11-09: "I have only seen capacitors with an "indention
    // ring" at the plus-side"): its groove is near the PLUS end - away from the marked minus lead - by the rubber seal
    // standing out of that end, and clear of the print.
    // ###########################################################################################
    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    public void An_axial_electrolytics_groove_is_by_its_plus_end_with_the_seal(int minus)
    {
        PartPictureGeometry.Layout layout = PartPictureGeometry.For(Picture(PartPicture.Shape.AxialElectrolytic, 2, minus));
        Rect groove = layout.Band!.Value;
        Rect seal = layout.Tab!.Value;

        Assert.Equal(minus == 0, groove.Center.X > layout.Body.Center.X);
        Assert.Equal(minus == 0, seal.Left > layout.Body.Center.X);
        Assert.True(layout.Body.Contains(groove));
        Assert.True(groove.Width < layout.Body.Width * 0.1, "a groove, not a band");
        Assert.False(layout.Marking.Intersects(groove));

        // Nearer its end than the middle.
        double plusEnd = minus == 0 ? layout.Body.Right : layout.Body.Left;
        Assert.True(Math.Abs(groove.Center.X - plusEnd) < layout.Body.Width * 0.25);
    }

    // A radial electrolytic's plus leg is the longer, as the part comes (owner request, 2026-11-09); the labels still
    // stand on one line, under the longer leg.
    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    public void A_radial_electrolytics_plus_leg_is_the_longer(int minus)
    {
        PartPictureGeometry.Layout layout = PartPictureGeometry.For(Picture(PartPicture.Shape.RadialElectrolytic, 2, minus));
        Rect minusLeg = layout.Leads[minus].Wire, plusLeg = layout.Leads[1 - minus].Wire;

        Assert.True(plusLeg.Bottom > minusLeg.Bottom + 8, $"plus {plusLeg} minus {minusLeg}");
        Assert.Equal(minusLeg.Top, plusLeg.Top, 3);
        Assert.Equal(layout.Leads[0].Label.Top, layout.Leads[1].Label.Top, 3);

        // Other parts standing keep their legs alike.
        PartPictureGeometry.Layout disc = PartPictureGeometry.For(Picture(PartPicture.Shape.DiscCapacitor, 2, minus));
        Assert.Equal(disc.Leads[0].Wire.Height, disc.Leads[1].Wire.Height, 3);
    }

    // A part standing: its legs at an even pitch, centred under the body, from under it downwards, each label under its
    // foot.
    [Theory]
    [MemberData(nameof(Every_shape))]
    public void A_part_standing_has_its_legs_evenly_under_its_body_and_a_label_under_each(PartPicture.Shape shape, int leads)
    {
        if (Lying.Contains(shape))
        {
            return;
        }

        PartPictureGeometry.Layout layout = PartPictureGeometry.For(Picture(shape, leads));
        List<double> centres = layout.Leads.Select(lead => lead.Wire.Center.X).ToList();

        Assert.False(layout.IsAxial);
        Assert.Equal(layout.Body.Center.X, (centres[0] + centres[^1]) / 2, 3);

        for (int i = 1; i < centres.Count; i++)
        {
            Assert.Equal(centres[1] - centres[0], centres[i] - centres[i - 1], 3);
        }

        foreach (PartPictureGeometry.LeadPlace lead in layout.Leads)
        {
            Assert.True(lead.Wire.Top <= Math.Max(layout.Body.Bottom, layout.Tab?.Bottom ?? 0), $"{shape}: a leg starts under its body");
            Assert.True(lead.Wire.Bottom > layout.Body.Bottom);
            Assert.Equal(lead.Wire.Center.X, lead.Label.Center.X, 3);
            Assert.True(lead.Label.Top >= lead.Wire.Bottom);
            Assert.True(lead.NameBox.Top >= lead.Label.Bottom);
        }

        // Neighbouring leads' areas meet and never overlap.
        for (int i = 1; i < layout.Leads.Count; i++)
        {
            Assert.Equal(layout.Leads[i - 1].Area.Right, layout.Leads[i].Area.Left, 3);
        }
    }

    // The minus stripe runs down the minus leg's side of the can, the print beside it.
    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    public void A_radial_electrolytics_stripe_is_on_its_minus_legs_side(int minus)
    {
        PartPictureGeometry.Layout layout = PartPictureGeometry.For(Picture(PartPicture.Shape.RadialElectrolytic, 2, minus));
        Rect stripe = layout.Band!.Value;

        Assert.Equal(minus == 0, stripe.Center.X < layout.Body.Center.X);
        Assert.True(layout.Body.Contains(stripe));
        Assert.False(layout.Marking.Intersects(stripe));
        Assert.True(layout.Marking.Width > 20);
    }

    // ###########################################################################################
    // A TO-92 AS IT LOOKS (owner request, 2026-11-09: "it does not look like a transistor"): the top of its case shows over
    // its face, the ellipse's middle on the face's top edge; its legs leave the case close together, bend out and end under
    // their labels, long - longer than the case is tall. Every other part's legs run straight down their wires.
    // ###########################################################################################
    [Fact]
    public void A_TO92s_top_shows_over_its_face_and_its_legs_bend_out_from_the_case()
    {
        PartPictureGeometry.Layout layout = PartPictureGeometry.For(Picture(PartPicture.Shape.TO92, 3));
        Rect cap = layout.Tab!.Value;

        Assert.Equal(layout.Body.Top, cap.Center.Y, 3);
        Assert.Equal(layout.Body.Width, cap.Width, 3);
        Assert.True(cap.Top >= 0);

        foreach (PartPictureGeometry.LeadPlace lead in layout.Leads)
        {
            IReadOnlyList<Point> path = PartPictureGeometry.LegPath(layout, lead);

            Assert.True(path[0].Y <= layout.Body.Bottom, "the leg comes out of the case");
            Assert.True(Math.Abs(path[0].X - layout.Body.Center.X) <= Math.Abs(lead.Wire.Center.X - layout.Body.Center.X), "closer together at the case");
            Assert.InRange(path[0].X, layout.Body.Left, layout.Body.Right);
            Assert.Equal(lead.Label.Center.X, path[^1].X, 3);
            Assert.Equal(lead.Wire.Bottom, path[^1].Y, 3);
            Assert.True(path[^1].Y - path[0].Y > layout.Body.Height, "the legs are long");

            // Downwards all the way.
            for (int i = 1; i < path.Count; i++)
            {
                Assert.True(path[i].Y >= path[i - 1].Y);
            }
        }

        // The middle leg goes straight down; a part of another shape too.
        IReadOnlyList<Point> middle = PartPictureGeometry.LegPath(layout, layout.Leads[1]);
        Assert.All(middle, point => Assert.Equal(layout.Leads[1].Wire.Center.X, point.X, 3));

        PartPictureGeometry.Layout disc = PartPictureGeometry.For(Picture(PartPicture.Shape.DiscCapacitor, 2));
        Assert.Equal(2, PartPictureGeometry.LegPath(disc, disc.Leads[0]).Count);
    }

    [Fact]
    public void A_TO220s_hole_is_in_its_tab_over_its_body()
    {
        PartPictureGeometry.Layout layout = PartPictureGeometry.For(Picture(PartPicture.Shape.TO220, 3));
        Rect tab = layout.Tab!.Value;

        Assert.True(tab.Contains(layout.Hole!.Value));
        Assert.True(tab.Bottom > layout.Body.Top && tab.Top < layout.Body.Top, "the tab stands over the body");
        Assert.Equal(tab.Width, layout.Body.Width);
    }

    [Fact]
    public void A_crystal_and_a_disc_stand_their_legs_under_their_base()
    {
        PartPictureGeometry.Layout crystal = PartPictureGeometry.For(Picture(PartPicture.Shape.Crystal, 3));
        Assert.True(crystal.Tab!.Value.Top < crystal.Body.Bottom && crystal.Tab.Value.Width > crystal.Body.Width, "the base is wider than the can");
        Assert.All(crystal.Leads, lead => Assert.True(lead.Wire.Top >= crystal.Body.Bottom));

        PartPictureGeometry.Layout disc = PartPictureGeometry.For(Picture(PartPicture.Shape.DiscCapacitor, 2));
        Rect neck = disc.Tab!.Value;
        Assert.All(disc.Leads, lead => Assert.True(neck.Left < lead.Wire.Left && lead.Wire.Right < neck.Right, "the coating runs down round each leg"));
    }

    // A resistor pack's body spans every leg, its dot at pin 1's end; a pin header's posts stand over its base, one per leg.
    [Fact]
    public void A_resistor_packs_body_spans_its_legs_and_its_dot_is_by_pin_1()
    {
        PartPictureGeometry.Layout pack = PartPictureGeometry.For(Picture(PartPicture.Shape.ResistorNetwork, 9, marked: 0));

        Assert.True(pack.Body.Left < pack.Leads[0].Wire.Left && pack.Body.Right > pack.Leads[^1].Wire.Right);
        Assert.True(pack.Dot!.Value.X < pack.Leads[1].Wire.Left);
        Assert.True(pack.Body.Contains(pack.Dot.Value));
        Assert.False(pack.Marking.Contains(pack.Dot.Value));

        PartPictureGeometry.Layout header = PartPictureGeometry.For(Picture(PartPicture.Shape.PinHeader, 4));

        Assert.All(header.Leads, lead =>
        {
            Assert.True(lead.Post!.Value.Bottom >= header.Body.Top && lead.Post.Value.Top < header.Body.Top, "a post over the base");
            Assert.Equal(lead.Wire.Center.X, lead.Post.Value.Center.X, 3);
        });
    }

    // Pointed at, a lead is found by its wire and its label; between two legs or off the picture, none.
    [Theory]
    [MemberData(nameof(Every_shape))]
    public void A_lead_is_found_where_it_is_pointed_at(PartPicture.Shape shape, int leads)
    {
        PartPictureGeometry.Layout layout = PartPictureGeometry.For(Picture(shape, leads));

        foreach (PartPictureGeometry.LeadPlace lead in layout.Leads)
        {
            Assert.Equal(lead.Number, PartPictureGeometry.LeadAt(layout, lead.Wire.Center)?.Number);
            Assert.Equal(lead.Number, PartPictureGeometry.LeadAt(layout, lead.Label.Center)?.Number);
        }

        Assert.Null(PartPictureGeometry.LeadAt(layout, new Point(-5, -5)));
        Assert.Null(PartPictureGeometry.LeadAt(layout, new Point(layout.Size.Width + 5, layout.Size.Height / 2)));
    }

    // ###########################################################################################
    // A name too long for its leg's pitch ("common" on a resistor pack) staggers every other name a line lower, so each
    // has two pitches' room and none is cut short or runs into the next; signs and letters never stagger.
    // ###########################################################################################
    [Fact]
    public void Names_too_long_for_their_pitch_stand_on_two_lines()
    {
        PartPicture.Picture pack = Picture(PartPicture.Shape.ResistorNetwork, 9, 0, "common", "R1", "R2", "R3", "R4", "R5", "R6", "R7", "R8");
        PartPictureGeometry.Layout layout = PartPictureGeometry.For(pack);

        Assert.True(PartPictureGeometry.NamesStagger(pack, layout.Leads[0].Label.Width));
        Assert.Equal(layout.Leads[0].NameBox.Top, layout.Leads[2].NameBox.Top);
        Assert.True(layout.Leads[1].NameBox.Top > layout.Leads[0].NameBox.Top);
        Assert.True(layout.Leads[0].NameBox.Width >= "common".Length * PartPictureGeometry.NameLetterWidth);

        // Names on one line never overlap.
        for (int i = 2; i < layout.Leads.Count; i++)
        {
            Assert.True(layout.Leads[i - 2].NameBox.Right <= layout.Leads[i].NameBox.Left);
        }

        Assert.False(PartPictureGeometry.NamesStagger(Picture(PartPicture.Shape.TO92, 3, null, "E", "C", "B"), 20));
        Assert.False(PartPictureGeometry.NamesStagger(Picture(PartPicture.Shape.RadialElectrolytic, 2, 1, "Positive", "Negative"), 20));

        PartPictureGeometry.Layout plain = PartPictureGeometry.For(Picture(PartPicture.Shape.TO92, 3, null, "E", "C", "B"));
        Assert.All(plain.Leads, lead => Assert.Equal(plain.Leads[0].NameBox.Top, lead.NameBox.Top));
    }
}
