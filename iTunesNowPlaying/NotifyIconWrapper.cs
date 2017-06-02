using System;
using System.ComponentModel;
using System.Windows;

namespace iTunesNowPlaying
{
    public partial class NotifyIconWrapper : Component
    {
        public NotifyIconWrapper()
        {
            InitializeComponent();

            // コンテキストメニューのイベントハンドラを登録
            this.toolStripMenuItem_Quit.Click += ToolStripMenuItem_Quit_Click;
        }

        public NotifyIconWrapper(IContainer container)
        {
            container.Add(this);

            InitializeComponent();
        }

        private void ToolStripMenuItem_Quit_Click(object sender, EventArgs e)
        {
            this.toolStripMenuItem_Quit.Click -= ToolStripMenuItem_Quit_Click;
            Application.Current.Shutdown();
        }
    }
}
