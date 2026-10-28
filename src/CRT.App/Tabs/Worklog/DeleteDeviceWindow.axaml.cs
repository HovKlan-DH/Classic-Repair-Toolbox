using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Interactivity;
using Handlers.DataHandling;

namespace CRT
{
    // ###########################################################################################
    // Confirms deleting one of the user's devices (CRT 4.0 step 4). Closes with true to delete, null
    // otherwise; Escape and Enter both cancel.
    // ###########################################################################################
    public partial class DeleteDeviceWindow : Window
    {
        public DeleteDeviceWindow()
        {
            this.InitializeComponent();

            this.AddHandler(KeyDownEvent, this.OnWindowPreviewKeyDown, RoutingStrategies.Tunnel);
        }

        public void Initialize(OwnedDeviceRecord device, int repairCount)
        {
            this.QuestionText.Text = OwnedDeviceWording.DeleteQuestion(device.Name, repairCount);
        }

        private void OnWindowPreviewKeyDown(object? sender, KeyEventArgs e)
        {
            if (e.Key == Key.Escape || e.Key == Key.Enter)
            {
                this.Close(null);
                e.Handled = true;
            }
        }

        private void OnCancelClick(object? sender, RoutedEventArgs e) => this.Close(null);

        private void OnDeleteClick(object? sender, RoutedEventArgs e) => this.Close(true);
    }
}
