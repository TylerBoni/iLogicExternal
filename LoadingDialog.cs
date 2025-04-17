using System;
using System.Drawing;
using System.Threading;
using System.Windows.Forms;

namespace iLogicExternal
{
    public partial class LoadingDialog : Form
    {
        private static LoadingDialog instance;
        private static readonly object lockObject = new object();
        private Timer updateTimer;
        private int dotCount = 0;

        public LoadingDialog()
        {
            InitializeComponent();

            // Set the form to be borderless
            this.FormBorderStyle = FormBorderStyle.None;
            this.StartPosition = FormStartPosition.CenterScreen;
            this.ShowInTaskbar = false;
            this.TopMost = true;

            // Set a semi-transparent background
            this.BackColor = Color.FromArgb(50, 50, 50);
            this.Opacity = 0.9;

            // Round the corners
            this.Region = System.Drawing.Region.FromHrgn(CreateRoundRectRgn(0, 0, Width, Height, 20, 20));

            // Create a timer to animate the dots
            updateTimer = new Timer();
            updateTimer.Interval = 400; // Update every 400ms
            updateTimer.Tick += UpdateTimer_Tick;
        }

        private void InitializeComponent()
        {
            this.SuspendLayout();

            // Set form properties
            this.AutoScaleDimensions = new System.Drawing.SizeF(8F, 16F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(400, 200);
            this.Name = "LoadingDialog";
            this.Text = "Loading";
            this.ResumeLayout(false);

            // Create and position the label
            messageLabel = new Label();
            messageLabel.AutoSize = false;
            messageLabel.TextAlign = ContentAlignment.MiddleCenter;
            messageLabel.Dock = DockStyle.Fill;
            messageLabel.Font = new Font("Segoe UI", 12F, FontStyle.Regular, GraphicsUnit.Point, 0);
            messageLabel.ForeColor = Color.White;
            messageLabel.Text = "Loading iLogic Rules...";
            this.Controls.Add(messageLabel);
        }

        private Label messageLabel;

        private void UpdateTimer_Tick(object sender, EventArgs e)
        {
            // Update dots animation
            dotCount = (dotCount + 1) % 4;
            string dots = new string('.', dotCount);
            messageLabel.Text = $"Loading iLogic Rules{dots}";
        }

        public void SetMessage(string message)
        {
            if (this.InvokeRequired)
            {
                this.Invoke(new Action<string>(SetMessage), message);
                return;
            }

            messageLabel.Text = message;
        }

        public static void ShowLoading(string message = "Loading iLogic Rules...")
        {
            try
            {
                // Ensure we're on the UI thread
                if (SynchronizationContext.Current == null)
                {
                    // If not on UI thread, we can't show a form
                    return;
                }

                // Create or get the instance
                lock (lockObject)
                {
                    if (instance == null || instance.IsDisposed)
                    {
                        instance = new LoadingDialog();
                    }

                    // Set the message and show the form
                    instance.SetMessage(message);

                    // Start the animation timer
                    instance.updateTimer.Start();

                    // Show the form if it's not already visible
                    if (!instance.Visible)
                    {
                        instance.Show();
                        instance.BringToFront();

                        // Refresh to make sure it's displayed immediately
                        instance.Refresh();
                        Application.DoEvents();
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error showing loading dialog: {ex.Message}");
            }
        }

        public static void HideLoading()
        {
            try
            {
                lock (lockObject)
                {
                    if (instance != null && !instance.IsDisposed)
                    {
                        // Stop the timer
                        instance.updateTimer.Stop();

                        // Hide the form
                        if (instance.InvokeRequired)
                        {
                            instance.Invoke(new Action(() =>
                            {
                                instance.Hide();
                            }));
                        }
                        else
                        {
                            instance.Hide();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                System.Diagnostics.Debug.WriteLine($"Error hiding loading dialog: {ex.Message}");
            }
        }

        protected override void OnPaint(PaintEventArgs e)
        {
            base.OnPaint(e);

            // Draw a border around the form
            using (Pen pen = new Pen(Color.FromArgb(80, 80, 80), 2))
            {
                e.Graphics.DrawRoundedRectangle(pen, 1, 1, Width - 3, Height - 3, 20);
            }
        }

        // P/Invoke method to create rounded rectangle region
        [System.Runtime.InteropServices.DllImport("Gdi32.dll", EntryPoint = "CreateRoundRectRgn")]
        private static extern IntPtr CreateRoundRectRgn(int nLeftRect, int nTopRect, int nRightRect, int nBottomRect, int nWidthEllipse, int nHeightEllipse);
    }

    // Extension method for drawing rounded rectangles
    public static class GraphicsExtensions
    {
        public static void DrawRoundedRectangle(this Graphics graphics, Pen pen, float x, float y, float width, float height, float radius)
        {
            RectangleF upperLeft = new RectangleF(x, y, radius * 2, radius * 2);
            RectangleF upperRight = new RectangleF(x + width - radius * 2, y, radius * 2, radius * 2);
            RectangleF lowerLeft = new RectangleF(x, y + height - radius * 2, radius * 2, radius * 2);
            RectangleF lowerRight = new RectangleF(x + width - radius * 2, y + height - radius * 2, radius * 2, radius * 2);

            // Draw arcs for corners
            graphics.DrawArc(pen, upperLeft, 180, 90);
            graphics.DrawArc(pen, upperRight, 270, 90);
            graphics.DrawArc(pen, lowerLeft, 90, 90);
            graphics.DrawArc(pen, lowerRight, 0, 90);

            // Draw lines for sides
            graphics.DrawLine(pen, x + radius, y, x + width - radius, y);
            graphics.DrawLine(pen, x, y + radius, x, y + height - radius);
            graphics.DrawLine(pen, x + width, y + radius, x + width, y + height - radius);
            graphics.DrawLine(pen, x + radius, y + height, x + width - radius, y + height);
        }
    }
}