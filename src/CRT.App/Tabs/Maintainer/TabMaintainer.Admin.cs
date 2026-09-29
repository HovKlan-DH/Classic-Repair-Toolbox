using System.Threading.Tasks;
using Avalonia.Controls;

namespace CRT
{
    // ###########################################################################################
    // THE "ADMIN" SCREEN (owner request, 2026-09-27): the administrator's tools on the left, the
    // chosen one on the right. Offered only to an administrator (the server's isAdministrator); the
    // server refuses everyone else anyway.
    //
    // "Set maintainers" was here until the owner moved it to the Systems screen the same day ("As
    // admin I should be allowed to set a system maintainer in the 'Systems' list - so that should
    // be moved from 'Admin' section") - see SystemView.Maintainers.cs. What is left is "Unused
    // files", whose list reads every workbook in the tree and takes seconds: it is read on choosing
    // it (the first time Admin is shown, since it is then the only thing to show), on switching
    // tree, and on its Refresh button - never again merely on coming back to Admin.
    // ###########################################################################################
    public partial class TabMaintainer
    {
        private async void OnAdminSelectionChanged(object? sender, SelectionChangedEventArgs e)
        {
            this.ApplyModeVisibility();

            object? chosen = this.FindControl<ListBox>("AdminList")?.SelectedItem;

            if (ReferenceEquals(chosen, this.FindControl<ListBoxItem>("UnusedFilesItem")))
                await this.UnusedFilesAdmin.LoadAsync();
        }

        // On showing Admin: the first time, "Unused files" is chosen (its SelectionChanged reads it);
        // after that, whatever was chosen stays as it was - see the header.
        private Task EnterAdminAsync()
        {
            if (this.FindControl<ListBox>("AdminList") is ListBox list && list.SelectedItem is null)
                list.SelectedItem = this.FindControl<ListBoxItem>("UnusedFilesItem");

            return Task.CompletedTask;
        }

        // ###########################################################################################
        // After a pool changed (on the Systems screen): the queue (assigning somebody may have been
        // prompted by a submission on screen, and the queue is filtered by pools) and the systems
        // (they count maintainers).
        // ###########################################################################################
        private async Task AfterPoolChangeAsync()
        {
            await this.RefreshQueueAsync(background: true);
            await this.RefreshSystemsAsync(background: true);
        }

        private void ClearAdmin()
        {
            if (this.FindControl<ListBox>("AdminList") is ListBox list)
                list.SelectedItem = null;

            this.UnusedFilesAdmin.Clear();
        }
    }
}
