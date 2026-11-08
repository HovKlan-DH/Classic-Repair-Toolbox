using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.Shapes;
using Avalonia.Layout;
using Avalonia.Media;
using Handlers.DataHandling;
using Handlers.Theming;
using System;

namespace CRT
{
    // ###########################################################################################
    // ONE CARD IN "My devices" (TabWorkbooks.Devices.cs builds its content) - a Border of the tab's card look
    // (Border.WorkbooksJobCard, so its style key stays Border's), and a row ListRowDrag can move while the devices are
    // sorted by hand (owner request, 2026-11-07: "Sort manually in list (drag up/down)"): the list's items ARE the
    // cards, each its own container, so the card moved is the card shown - nothing is built again under a drag.
    //
    // While dragged it is the dashed slot every list of CRT shows where a drop will land, its own height.
    // ###########################################################################################
    internal sealed class DeviceListCard : Border, IDraggableRow
    {
        private readonly Control thisContent;

        private Panel? thisPlaceholder;

        private bool thisIsDropPlaceholder;

        public DeviceListCard(OwnedDeviceRecord device, Control content)
        {
            this.Device = device ?? throw new ArgumentNullException(nameof(device));
            this.thisContent = content ?? throw new ArgumentNullException(nameof(content));
            this.Tag = device;
            this.Child = content;
            this.Classes.Add("WorkbooksJobCard");

            // ListRowDrag finds the row it is pressed on by the handle's DataContext.
            this.DataContext = this;
        }

        // Styled as the Border it is - a selector names exact types.
        protected override Type StyleKeyOverride => typeof(Border);

        public OwnedDeviceRecord Device { get; }

        public double PlaceholderHeight { get; set; }

        public bool IsDropPlaceholder
        {
            get => this.thisIsDropPlaceholder;
            set
            {
                if (this.thisIsDropPlaceholder == value)
                {
                    return;
                }

                this.thisIsDropPlaceholder = value;

                if (value)
                {
                    this.thisPlaceholder ??= BuildPlaceholder();
                    this.thisPlaceholder.Height = this.PlaceholderHeight > 0 ? this.PlaceholderHeight : double.NaN;
                    this.Classes.Remove("WorkbooksJobCard");
                    this.Padding = default;
                    this.BorderThickness = default;
                    this.Background = Brushes.Transparent;
                    this.Child = this.thisPlaceholder;
                }
                else
                {
                    this.ClearValue(PaddingProperty);
                    this.ClearValue(BorderThicknessProperty);
                    this.ClearValue(BackgroundProperty);
                    this.Classes.Add("WorkbooksJobCard");
                    this.Child = this.thisContent;
                }
            }
        }

        // The dashed slot - the look of every placeholder ListRowDrag shows (ComponentPane.axaml's rows).
        private static Panel BuildPlaceholder()
        {
            var slot = new Panel();

            slot.Children.Add(new Rectangle
            {
                Stroke = ThemeResources.ResolveBrush("Text_Fail_Fg", Brushes.IndianRed),
                StrokeThickness = 2,
                StrokeDashArray = [4, 3],
                RadiusX = 3,
                RadiusY = 3,
                Fill = ThemeResources.ResolveBrush("Bg", Brushes.Transparent),
            });

            slot.Children.Add(new TextBlock
            {
                Text = "Release to place the device here",
                FontSize = 10,
                Opacity = 0.7,
                TextWrapping = TextWrapping.Wrap,
                TextAlignment = TextAlignment.Center,
                HorizontalAlignment = HorizontalAlignment.Center,
                VerticalAlignment = VerticalAlignment.Center,
                Margin = new Thickness(6, 0),
            });

            return slot;
        }
    }
}
