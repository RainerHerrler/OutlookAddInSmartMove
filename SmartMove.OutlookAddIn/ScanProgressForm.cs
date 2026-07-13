using System;
using System.Drawing;
using System.Windows.Forms;

namespace SmartMove.OutlookAddIn
{
    internal sealed class ScanProgressForm : Form
    {
        private readonly Label statusLabel;
        private readonly Label countLabel;
        private readonly ProgressBar progressBar;
        private readonly Button cancelButton;
        private bool allowClose;

        public event EventHandler CancelRequested;

        public ScanProgressForm()
        {
            Text = "SmartMove – Statistik erstellen";
            ClientSize = new Size(600, 155);
            FormBorderStyle = FormBorderStyle.FixedDialog;
            MaximizeBox = false;
            MinimizeBox = false;
            ShowInTaskbar = false;
            StartPosition = FormStartPosition.CenterScreen;

            statusLabel = new Label
            {
                AutoEllipsis = true,
                Location = new Point(18, 18),
                Size = new Size(564, 22),
                Text = "Ordner werden ermittelt …"
            };
            countLabel = new Label
            {
                Location = new Point(18, 48),
                Size = new Size(564, 20),
                Text = "0 E-Mails ausgewertet"
            };
            progressBar = new ProgressBar
            {
                Location = new Point(18, 76),
                Size = new Size(564, 20),
                Minimum = 0,
                Maximum = 1
            };
            cancelButton = new Button
            {
                Location = new Point(477, 111),
                Size = new Size(105, 28),
                Text = "Abbrechen"
            };
            cancelButton.Click += (sender, args) => RequestCancel();

            Controls.Add(statusLabel);
            Controls.Add(countLabel);
            Controls.Add(progressBar);
            Controls.Add(cancelButton);
            FormClosing += OnFormClosing;
        }

        public void SetProgress(string folderPath, int completedFolders, int totalFolders, long messages, long matchedEmails)
        {
            statusLabel.Text = folderPath;
            countLabel.Text = string.Format(
                "{0:N0} Elemente geprüft, {1:N0} E-Mails gezählt – Ordner {2:N0} von {3:N0}",
                messages,
                matchedEmails,
                Math.Min(completedFolders + 1, totalFolders),
                totalFolders);
            progressBar.Maximum = Math.Max(1, totalFolders);
            progressBar.Value = Math.Max(0, Math.Min(completedFolders, progressBar.Maximum));
        }

        public void CloseAfterScan()
        {
            allowClose = true;
            Close();
        }

        private void RequestCancel()
        {
            cancelButton.Enabled = false;
            cancelButton.Text = "Wird beendet …";
            CancelRequested?.Invoke(this, EventArgs.Empty);
        }

        private void OnFormClosing(object sender, FormClosingEventArgs args)
        {
            if (allowClose)
            {
                return;
            }

            args.Cancel = true;
            RequestCancel();
        }
    }
}
