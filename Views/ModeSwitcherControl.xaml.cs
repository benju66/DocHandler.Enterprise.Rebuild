using System;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Media.Animation;
using Serilog;

namespace DocHandler.Views
{
    /// <summary>
    /// Mode switcher control for seamless mode transitions (Phase 2 Milestone 2 - Day 5)
    /// </summary>
    public partial class ModeSwitcherControl : UserControl
    {
        private readonly ILogger _logger;
        private bool _isHovering = false;

        public ModeSwitcherControl()
        {
            InitializeComponent();
            _logger = Log.ForContext<ModeSwitcherControl>();
            
            InitializeEvents();
            _logger.Debug("ModeSwitcherControl initialized");
        }

        private void InitializeEvents()
        {
            // Show mode selector on hover
            ModeIndicator.MouseEnter += (s, e) =>
            {
                _isHovering = true;
                ShowModeSelector();
            };

            ModeIndicator.MouseLeave += (s, e) =>
            {
                _isHovering = false;
                // Delay hiding to allow for interaction
                Dispatcher.BeginInvoke(new Action(() =>
                {
                    if (!_isHovering && !ModeSelectionPopup.IsMouseOver)
                    {
                        HideModeSelector();
                    }
                }), System.Windows.Threading.DispatcherPriority.Background);
            };

            // Keep popup open when hovering over it
            ModeSelectionPopup.MouseEnter += (s, e) => _isHovering = true;
            ModeSelectionPopup.MouseLeave += (s, e) =>
            {
                _isHovering = false;
                HideModeSelector();
            };
        }

        private void ShowModeSelector()
        {
            try
            {
                // Animate the mode selector button into view
                var fadeIn = new DoubleAnimation(0, 1, TimeSpan.FromMilliseconds(200));
                ModeSelectorButton.Visibility = Visibility.Visible;
                ModeSelectorButton.BeginAnimation(OpacityProperty, fadeIn);

                _logger.Debug("Mode selector shown");
            }
            catch (Exception ex)
            {
                _logger.Warning(ex, "Error showing mode selector");
            }
        }

        private void HideModeSelector()
        {
            try
            {
                if (!_isHovering && !ModeSelectionPopup.IsOpen)
                {
                    // Animate the mode selector button out of view
                    var fadeOut = new DoubleAnimation(1, 0, TimeSpan.FromMilliseconds(200));
                    fadeOut.Completed += (s, e) =>
                    {
                        if (!_isHovering)
                        {
                            ModeSelectorButton.Visibility = Visibility.Collapsed;
                        }
                    };
                    ModeSelectorButton.BeginAnimation(OpacityProperty, fadeOut);

                    _logger.Debug("Mode selector hidden");
                }
            }
            catch (Exception ex)
            {
                _logger.Warning(ex, "Error hiding mode selector");
            }
        }

        private void ModeSelectorButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // Toggle the popup
                ModeSelectionPopup.IsOpen = !ModeSelectionPopup.IsOpen;
                
                if (ModeSelectionPopup.IsOpen)
                {
                    _logger.Debug("Mode selection popup opened");
                }
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Error toggling mode selection popup");
            }
        }

        private void ModeButton_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                // Close the popup when a mode is selected
                ModeSelectionPopup.IsOpen = false;
                
                // Trigger mode change animation
                TriggerModeChangeAnimation();
                
                _logger.Debug("Mode button clicked, popup closed");
            }
            catch (Exception ex)
            {
                _logger.Error(ex, "Error handling mode button click");
            }
        }

        private void TriggerModeChangeAnimation()
        {
            try
            {
                // Find and trigger the mode change animation
                var storyboard = Resources["ModeChangeAnimation"] as Storyboard;
                if (storyboard != null)
                {
                    storyboard.Begin();
                    _logger.Debug("Mode change animation triggered");
                }
            }
            catch (Exception ex)
            {
                _logger.Warning(ex, "Error triggering mode change animation");
            }
        }

        /// <summary>
        /// Public method to trigger mode change animation from external code
        /// </summary>
        public void AnimateModeChange()
        {
            TriggerModeChangeAnimation();
        }

        /// <summary>
        /// Update the mode icon based on current mode
        /// </summary>
        public void UpdateModeIcon(string mode)
        {
            try
            {
                var symbol = mode switch
                {
                    "SaveQuotes" => ModernWpf.Controls.Symbol.Save,
                    "default" => ModernWpf.Controls.Symbol.Document,
                    _ => ModernWpf.Controls.Symbol.Setting
                };

                ModeIcon.Symbol = symbol;
                _logger.Debug("Mode icon updated for mode: {Mode}", mode);
            }
            catch (Exception ex)
            {
                _logger.Warning(ex, "Error updating mode icon for mode: {Mode}", mode);
            }
        }
    }
} 