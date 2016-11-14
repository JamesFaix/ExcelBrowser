using System;
using System.Windows;
using ExcelBrowser.Controller;
using ExcelBrowser.ViewModels;

namespace ExcelBrowser.View {
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window, IDisposable {

        public MainWindow() {// SessionMonitor monitor) {
                             //    Requires.NotNull(monitor, nameof(monitor));
            InitializeComponent();

            var vm = new SessionViewModel {
                Apps = {
                    new AppViewModel { ProcessId = 1234,
                        Books = {
                            new BookViewModel { Name = "Book1",
                                Sheets = {
                                    new SheetViewModel { Name = "Sheet1", Windows = { new SheetWindowViewModel(), new SheetWindowViewModel() } },
                                    new SheetViewModel { Name = "Sheet2", Windows = { new SheetWindowViewModel(), new SheetWindowViewModel() } },
                                    new SheetViewModel { Name = "Sheet3", Windows = { new SheetWindowViewModel(), new SheetWindowViewModel() } },
                                },
                                Windows = new BookWindowsViewModel {
                                    Windows = {
                                        new BookWindowViewModel { Index = 1 },
                                        new BookWindowViewModel { Index = 2 }
                                    }
                                }
                            },
                            new BookViewModel { Name = "Book2",
                                Sheets = {
                                    new SheetViewModel { Name = "Sheet1", Windows = { new SheetWindowViewModel() } },
                                    new SheetViewModel { Name = "Sheet2", Windows = { new SheetWindowViewModel() } },
                                    new SheetViewModel { Name = "Sheet3", Windows = { new SheetWindowViewModel() } },
                                },
                                Windows = new BookWindowsViewModel {
                                    Windows = {
                                        new BookWindowViewModel { Index = 1 }
                                    }
                                }
                            }
                        }
                    },
                    new AppViewModel { ProcessId = 2345,
                        Books = {
                            new BookViewModel { Name = "Book3",
                                Sheets = {
                                    new SheetViewModel { Name = "Sheet1", Windows = { new SheetWindowViewModel() }  },
                                    new SheetViewModel { Name = "Sheet2", Windows = { new SheetWindowViewModel() }  },
                                    new SheetViewModel { Name = "Sheet3", Windows = { new SheetWindowViewModel() }  },
                                },
                                Windows = new BookWindowsViewModel {
                                    Windows = {
                                        new BookWindowViewModel { Index = 1 }
                                    }
                                }
                            },
                            new BookViewModel() { Name = "Book4",
                                Sheets = {
                                    new SheetViewModel { Name = "Sheet1", Windows = { new SheetWindowViewModel() } },
                                    new SheetViewModel { Name = "Sheet2", Windows = { new SheetWindowViewModel() } },
                                    new SheetViewModel { Name = "Sheet3", Windows = { new SheetWindowViewModel() } },
                                },
                                Windows = new BookWindowsViewModel {
                                    Windows = {
                                        new BookWindowViewModel { Index = 1 }
                                    }
                                }
                            }
                        }
                    }
                }
            };

            ctrl_Session.ViewModel = vm;

            // this.monitor = monitor;            
        }

        //  private readonly SessionMonitor monitor;

        public void Dispose() {
            //   monitor.Dispose();
        }
    }
}
