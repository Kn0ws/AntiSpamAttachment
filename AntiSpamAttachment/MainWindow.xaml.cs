using System;
using System.Reflection;
using System.Windows;


namespace AntiSpamAttachment
{
    /// <summary>
    /// MainWindow.xaml の相互作用ロジック
    /// </summary>
    public partial class MainWindow : Window
    {
        public double version;
        
        public MainWindow()
        {
            InitializeComponent();
            // Excelでバージョンチェック
            double version = checkOffice("Excel.Application");
            bool isEmptyRegistry = LoadRegistry(version);
            if(isEmptyRegistry)
            {
                MessageBox.Show("既にブロックする拡張子が設定済みです。\r\nその他の拡張子を設定する場合は続行してください。");
            }
            else
            {
                MessageBox.Show("ブロックする拡張子を入力して「実行」ボタンを押してください。");
            }
        }

        private double checkOffice(string applicationID)
        {
            var type = Type.GetTypeFromProgID(applicationID);
            var application = Activator.CreateInstance(type);
            if(application == null)
            {
                MessageBox.Show("Officeがインストールされていないか、検出出来ませんでした。\r\nシステムを終了します。");
                Application.Current.Shutdown();
                return 0;
            }
            else
            {
                var ver = application.GetType().InvokeMember("Version", BindingFlags.GetProperty, null, application, null);
                Double.TryParse(ver.ToString(), out version);
                if (version >= 14.0)
                {
                    // Office2010以降がインストール済みの場合
                    return version;
                }
                else
                {
                    // Office 2000より前のバージョン
                    MessageBox.Show("本ソフトウェアのサポート対象外のOfficeです。システムを終了します。");
                    Application.Current.Shutdown();
                    return 0;
                }
            }
        }

        private bool LoadRegistry(double version)
        {

            Microsoft.Win32.RegistryKey regkey = Microsoft.Win32.Registry.CurrentUser.OpenSubKey(@"Software\Microsoft\Office\" + version + @".0\Outlook\Security\", false);
            string check = (string)regkey.GetValue("Level1Add");
            regkey.Close();
            if(check != null)
            {
                extensions.Text = check;
                return true;
            }
            else
            {
                return false;
            }
            
        }

        private void submit_Click(object sender, RoutedEventArgs e)
        {
            string extension = extensions.Text;
            Microsoft.Win32.RegistryKey regkey = Microsoft.Win32.Registry.CurrentUser.CreateSubKey(@"Software\Microsoft\Office\" + version + @".0\Outlook\Security\");
            regkey.SetValue("Level1Add", extension);
            regkey.Close();
            MessageBox.Show("変更を適用しました。\r\n この機能は次回Outlookを再起動した際に有効になります。\r\n現在Outlookが起動している場合は一度終了してから再度起動してください。");

        }

        
    }

}
