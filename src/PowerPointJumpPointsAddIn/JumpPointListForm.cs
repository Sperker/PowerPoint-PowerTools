using System;
using System.Collections.Generic;
using System.Drawing;
using System.Runtime.InteropServices;
using System.Windows.Forms;

namespace PowerPointJumpPointsAddIn
{
    public sealed class JumpPointListForm : Form
    {
        private static readonly IntPtr HWND_TOPMOST = new IntPtr(-1);
        private static readonly IntPtr HWND_NOTOPMOST = new IntPtr(-2);
        private const uint SWP_NOSIZE = 0x0001;
        private const uint SWP_NOMOVE = 0x0002;
        private const uint SWP_NOACTIVATE = 0x0010;
        private const uint TOPMOST_FLAGS = SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE;

        [DllImport("user32.dll")]
        private static extern bool SetWindowPos(
            IntPtr hWnd,
            IntPtr hWndInsertAfter,
            int x,
            int y,
            int cx,
            int cy,
            uint uFlags);

        private readonly JumpPointController _controller;

        private readonly ListView _listView;
        private readonly Button _previousJumpPointButton;
        private readonly Button _nextJumpPointButton;
        private readonly Button _goToButton;
        private readonly Button _renameButton;
        private readonly Button _deleteButton;
        private readonly Button _deleteAllButton;
        private readonly Button _refreshButton;
        private readonly Button _closeButton;
        private readonly Label _summaryLabel;
        private readonly Label _navigationLabel;
        private readonly Label _actionsLabel;
        private readonly CheckBox _topMostCheckBox;

        public JumpPointListForm(JumpPointController controller)
        {
            _controller = controller;

            Text = "Jump Points Navigator";
            StartPosition = FormStartPosition.CenterScreen;
            FormBorderStyle = FormBorderStyle.Sizable;
            MinimizeBox = true;
            ShowInTaskbar = false;
            MinimumSize = new Size(455, 540);
            ClientSize = new Size(560, 580);

            Label titleLabel = new Label
            {
                Text = "Jump Points in current presentation",
                AutoSize = true,
                Font = new Font(Font.FontFamily, 10, FontStyle.Bold),
                Location = new Point(12, 12)
            };

            _topMostCheckBox = new CheckBox
            {
                Text = "Keep this window on top",
                AutoSize = true,
                Location = new Point(355, 13),
                Anchor = AnchorStyles.Top | AnchorStyles.Right,
                Checked = true
            };
            _topMostCheckBox.CheckedChanged += TopMostCheckBox_CheckedChanged;

            _summaryLabel = new Label
            {
                Text = string.Empty,
                AutoSize = false,
                Location = new Point(12, 38),
                Size = new Size(520, 22),
                Anchor = AnchorStyles.Top | AnchorStyles.Left | AnchorStyles.Right
            };

            _listView = new ListView
            {
                Location = new Point(12, 66),
                Size = new Size(536, 370),
                View = View.Details,
                FullRowSelect = true,
                GridLines = true,
                MultiSelect = true,
                HideSelection = false,
                Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right
            };

            _listView.Columns.Add("Slide", 58);
            _listView.Columns.Add("Slide ID", 76);
            _listView.Columns.Add("Jump Point Name", 210);
            _listView.Columns.Add("Detected Title / First text", 260);
            _listView.Columns.Add("Created", 145);
            _listView.Columns.Add("Version", 70);
            _listView.DoubleClick += ListView_DoubleClick;

            _navigationLabel = new Label
            {
                Text = "Jump Point navigation:",
                AutoSize = true,
                Location = new Point(12, 452),
                Anchor = AnchorStyles.Bottom | AnchorStyles.Left
            };

            _previousJumpPointButton = new Button
            {
                Text = "Previous Jump Point",
                Location = new Point(150, 446),
                Size = new Size(145, 28),
                Anchor = AnchorStyles.Bottom | AnchorStyles.Left
            };
            _previousJumpPointButton.Click += PreviousJumpPointButton_Click;

            _nextJumpPointButton = new Button
            {
                Text = "Next Jump Point",
                Location = new Point(301, 446),
                Size = new Size(125, 28),
                Anchor = AnchorStyles.Bottom | AnchorStyles.Left
            };
            _nextJumpPointButton.Click += NextJumpPointButton_Click;

            _actionsLabel = new Label
            {
                Text = "Navigator actions:",
                AutoSize = true,
                Location = new Point(12, 494),
                Anchor = AnchorStyles.Bottom | AnchorStyles.Left
            };

            _goToButton = new Button
            {
                Text = "Go To",
                Location = new Point(132, 488),
                Size = new Size(67, 28),
                Anchor = AnchorStyles.Bottom | AnchorStyles.Left
            };
            _goToButton.Click += GoToButton_Click;

            _renameButton = new Button
            {
                Text = "Rename",
                Location = new Point(204, 488),
                Size = new Size(73, 28),
                Anchor = AnchorStyles.Bottom | AnchorStyles.Left
            };
            _renameButton.Click += RenameButton_Click;

            _deleteButton = new Button
            {
                Text = "Delete",
                Location = new Point(282, 488),
                Size = new Size(67, 28),
                Anchor = AnchorStyles.Bottom | AnchorStyles.Left
            };
            _deleteButton.Click += DeleteButton_Click;

            _deleteAllButton = new Button
            {
                Text = "Delete All",
                Location = new Point(354, 488),
                Size = new Size(82, 28),
                Anchor = AnchorStyles.Bottom | AnchorStyles.Left
            };
            _deleteAllButton.Click += DeleteAllButton_Click;

            _refreshButton = new Button
            {
                Text = "Refresh List",
                Location = new Point(441, 488),
                Size = new Size(95, 28),
                Anchor = AnchorStyles.Bottom | AnchorStyles.Left
            };
            _refreshButton.Click += RefreshButton_Click;

            _closeButton = new Button
            {
                Text = "Close",
                Location = new Point(460, 534),
                Size = new Size(80, 28),
                Anchor = AnchorStyles.Bottom | AnchorStyles.Right
            };
            _closeButton.Click += CloseButton_Click;

            Controls.Add(titleLabel);
            Controls.Add(_topMostCheckBox);
            Controls.Add(_summaryLabel);
            Controls.Add(_listView);
            Controls.Add(_navigationLabel);
            Controls.Add(_previousJumpPointButton);
            Controls.Add(_nextJumpPointButton);
            Controls.Add(_actionsLabel);
            Controls.Add(_goToButton);
            Controls.Add(_renameButton);
            Controls.Add(_deleteButton);
            Controls.Add(_deleteAllButton);
            Controls.Add(_refreshButton);
            Controls.Add(_closeButton);

            Load += JumpPointListForm_Load;
            Activated += JumpPointListForm_Activated;
            Shown += JumpPointListForm_Shown;
        }

        private void JumpPointListForm_Load(object sender, EventArgs e)
        {
            RefreshJumpPointList();
        }

        private void JumpPointListForm_Shown(object sender, EventArgs e)
        {
            ApplyTopMostState();
        }

        private void JumpPointListForm_Activated(object sender, EventArgs e)
        {
            ApplyTopMostState();
        }

        public void RefreshJumpPointList()
        {
            int? previouslySelectedSlideId = GetFirstSelectedSlideId();
            _listView.Items.Clear();

            List<JumpPointListEntry> entries;

            try
            {
                entries = _controller.GetJumpPointListEntries();
            }
            catch
            {
                entries = new List<JumpPointListEntry>();
            }

            foreach (JumpPointListEntry entry in entries)
            {
                ListViewItem item = new ListViewItem(entry.SlideIndex.ToString());
                item.SubItems.Add(entry.SlideId.ToString());
                item.SubItems.Add(string.IsNullOrWhiteSpace(entry.Label) ? "-" : entry.Label);
                item.SubItems.Add(string.IsNullOrWhiteSpace(entry.Title) ? "(no title/text found)" : entry.Title);
                item.SubItems.Add(string.IsNullOrWhiteSpace(entry.Created) ? "-" : entry.Created);
                item.SubItems.Add(string.IsNullOrWhiteSpace(entry.Version) ? "-" : entry.Version);
                item.Tag = entry.SlideId;

                if (previouslySelectedSlideId.HasValue && previouslySelectedSlideId.Value == entry.SlideId)
                {
                    item.Selected = true;
                    item.Focused = true;
                }

                _listView.Items.Add(item);
            }

            _summaryLabel.Text = entries.Count == 1
                ? "1 jump point found."
                : entries.Count + " jump points found.";

            if (_listView.SelectedItems.Count > 0)
            {
                _listView.SelectedItems[0].EnsureVisible();
            }

            ApplyTopMostState();
        }

        public void HighlightSlideById(int slideId)
        {
            try
            {
                bool found = false;

                foreach (ListViewItem item in _listView.Items)
                {
                    bool isCurrentSlide = item.Tag is int && (int)item.Tag == slideId;
                    item.Selected = isCurrentSlide;
                    item.Focused = isCurrentSlide;

                    if (isCurrentSlide)
                    {
                        item.EnsureVisible();
                        found = true;
                    }
                }

                if (!found)
                {
                    _listView.SelectedItems.Clear();
                }

                ApplyTopMostState();
            }
            catch
            {
                // Ignore highlight issues.
            }
        }

        public void ClearCurrentSlideSelection()
        {
            try
            {
                _listView.SelectedItems.Clear();
                ApplyTopMostState();
            }
            catch
            {
                // Ignore selection issues.
            }
        }

        private int? GetFirstSelectedSlideId()
        {
            if (_listView.SelectedItems.Count == 0)
            {
                return null;
            }

            object tag = _listView.SelectedItems[0].Tag;

            if (tag is int)
            {
                return (int)tag;
            }

            return null;
        }

        private List<int> GetSelectedSlideIds()
        {
            List<int> result = new List<int>();

            foreach (ListViewItem item in _listView.SelectedItems)
            {
                if (item.Tag is int)
                {
                    result.Add((int)item.Tag);
                }
            }

            return result;
        }

        private string GetFirstSelectedCurrentLabel()
        {
            try
            {
                if (_listView.SelectedItems.Count == 0)
                {
                    return string.Empty;
                }

                string value = _listView.SelectedItems[0].SubItems[2].Text;

                if (value == "-")
                {
                    return string.Empty;
                }

                return value;
            }
            catch
            {
                return string.Empty;
            }
        }

        private int GetFirstSelectedSlideIndex()
        {
            try
            {
                if (_listView.SelectedItems.Count == 0)
                {
                    return 0;
                }

                int slideIndex;
                if (int.TryParse(_listView.SelectedItems[0].Text, out slideIndex))
                {
                    return slideIndex;
                }
            }
            catch
            {
                // Ignore.
            }

            return 0;
        }

        private void GoToSelectedSlide()
        {
            int? slideId = GetFirstSelectedSlideId();

            if (!slideId.HasValue)
            {
                MessageBox.Show(
                    "Please select a jump point first.",
                    AddInInfo.ProductName,
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);

                return;
            }

            bool success = _controller.NavigateToSlideById(slideId.Value);

            if (!success)
            {
                MessageBox.Show(
                    "Could not navigate to the selected jump point.",
                    AddInInfo.ProductName,
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Warning);

                return;
            }

            RefreshJumpPointList();
            HighlightSlideById(slideId.Value);
            ApplyTopMostState();
        }

        private void GoToPreviousJumpPointFromNavigator()
        {
            _controller.GoToPreviousJumpPoint();
            RefreshJumpPointList();
            _controller.NotifyCurrentSlideChanged();
            ApplyTopMostState();
        }

        private void GoToNextJumpPointFromNavigator()
        {
            _controller.GoToNextJumpPoint();
            RefreshJumpPointList();
            _controller.NotifyCurrentSlideChanged();
            ApplyTopMostState();
        }

        private void RenameSelectedJumpPoint()
        {
            int? slideId = GetFirstSelectedSlideId();

            if (!slideId.HasValue)
            {
                MessageBox.Show(
                    "Please select a jump point to rename.",
                    AddInInfo.ProductName,
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);

                return;
            }

            using (JumpPointRenameForm form = new JumpPointRenameForm(GetFirstSelectedCurrentLabel(), GetFirstSelectedSlideIndex()))
            {
                DialogResult result = form.ShowDialog(this);

                if (result != DialogResult.OK)
                {
                    return;
                }

                bool success = _controller.SetJumpPointLabelBySlideId(slideId.Value, form.JumpPointName);

                if (!success)
                {
                    MessageBox.Show(
                        "Could not rename the selected jump point.",
                        AddInInfo.ProductName,
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning);

                    return;
                }

                RefreshJumpPointList();
                HighlightSlideById(slideId.Value);
            }
        }

        private void DeleteSelectedJumpPoints()
        {
            List<int> selectedSlideIds = GetSelectedSlideIds();

            if (selectedSlideIds.Count == 0)
            {
                MessageBox.Show(
                    "Please select one or more jump points to delete.",
                    AddInInfo.ProductName,
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);

                return;
            }

            if (_controller.IsNavigatorConfirmationEnabled())
            {
                DialogResult result = MessageBox.Show(
                    selectedSlideIds.Count == 1
                        ? "Remove the selected jump point?"
                        : "Remove the selected " + selectedSlideIds.Count + " jump points?",
                    AddInInfo.ProductName,
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Warning);

                if (result != DialogResult.Yes)
                {
                    return;
                }
            }

            int removedCount = 0;

            foreach (int slideId in selectedSlideIds)
            {
                if (_controller.RemoveJumpPointBySlideId(slideId))
                {
                    removedCount++;
                }
            }

            RefreshJumpPointList();

            if (_controller.IsNavigatorConfirmationEnabled())
            {
                MessageBox.Show(
                    removedCount == 1
                        ? "Removed 1 jump point."
                        : "Removed " + removedCount + " jump points.",
                    AddInInfo.ProductName,
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
            }
        }

        private void DeleteAllJumpPointsFromNavigator()
        {
            int count = _controller.GetJumpPointCount();

            if (count == 0)
            {
                MessageBox.Show(
                    "No jump points defined in this presentation.",
                    AddInInfo.ProductName,
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);

                return;
            }

            if (_controller.IsNavigatorConfirmationEnabled())
            {
                DialogResult result = MessageBox.Show(
                    "Remove all " + count + " jump point(s) from this presentation?",
                    AddInInfo.ProductName,
                    MessageBoxButtons.YesNo,
                    MessageBoxIcon.Warning);

                if (result != DialogResult.Yes)
                {
                    return;
                }
            }

            int removedCount = _controller.ClearAllJumpPointsFromNavigator();
            RefreshJumpPointList();

            if (_controller.IsNavigatorConfirmationEnabled())
            {
                MessageBox.Show(
                    "Removed " + removedCount + " jump point(s).",
                    AddInInfo.ProductName,
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information);
            }
        }

        private void ApplyTopMostState()
        {
            try
            {
                TopMost = _topMostCheckBox.Checked;

                if (IsHandleCreated)
                {
                    SetWindowPos(
                        Handle,
                        _topMostCheckBox.Checked ? HWND_TOPMOST : HWND_NOTOPMOST,
                        0,
                        0,
                        0,
                        0,
                        TOPMOST_FLAGS);
                }
            }
            catch
            {
                // Ignore top-most issues.
            }
        }

        private void TopMostCheckBox_CheckedChanged(object sender, EventArgs e)
        {
            ApplyTopMostState();
        }

        private void ListView_DoubleClick(object sender, EventArgs e)
        {
            GoToSelectedSlide();
        }

        private void PreviousJumpPointButton_Click(object sender, EventArgs e)
        {
            GoToPreviousJumpPointFromNavigator();
        }

        private void NextJumpPointButton_Click(object sender, EventArgs e)
        {
            GoToNextJumpPointFromNavigator();
        }

        private void GoToButton_Click(object sender, EventArgs e)
        {
            GoToSelectedSlide();
        }

        private void RenameButton_Click(object sender, EventArgs e)
        {
            RenameSelectedJumpPoint();
        }

        private void DeleteButton_Click(object sender, EventArgs e)
        {
            DeleteSelectedJumpPoints();
        }

        private void DeleteAllButton_Click(object sender, EventArgs e)
        {
            DeleteAllJumpPointsFromNavigator();
        }

        private void RefreshButton_Click(object sender, EventArgs e)
        {
            RefreshJumpPointList();
        }

        private void CloseButton_Click(object sender, EventArgs e)
        {
            Close();
        }
    }
}
