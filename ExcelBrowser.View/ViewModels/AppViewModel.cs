﻿using System.Collections.ObjectModel;
using System.Windows;
using System.Windows.Media;
using ExcelBrowser.Model;

namespace ExcelBrowser.ViewModels {

    public class AppViewModel {

        public AppViewModel(AppToken token) {
            Requires.NotNull(token, nameof(token));
            Token = token;
            Books = new ObservableCollection<BookViewModel>();
        }

        public AppToken Token { get; }

        public string Label => $"{Token.Version} [ProcessID: {Token.Id.ProcessId}]";

        public ObservableCollection<BookViewModel> Books { get; set; }
        
        public Brush Foreground => new SolidColorBrush { Color = Colors.White };
        public Brush Background => new SolidColorBrush { Color = Colors.DarkGreen };
        public FontWeight FontWeight => Token.IsActive ? FontWeights.Bold : FontWeights.Normal;
    }
}
