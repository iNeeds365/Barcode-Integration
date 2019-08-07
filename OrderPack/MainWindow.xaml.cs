using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows;
using System.Windows.Controls;
using System.Windows.Data;
using System.Windows.Documents;
using System.Windows.Input;
using System.Windows.Media;
using System.Windows.Media.Imaging;
using System.Windows.Navigation;
using System.Windows.Shapes;
using System.Windows.Forms;
using System.Drawing;
using System.IO;
using System.Runtime.InteropServices;
using XlsFile;
using System.Windows.Threading;
using System.Threading;
using System.Media;

namespace OrderPack
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
    public partial class MainWindow : Window
    {
        public static readonly log4net.ILog logger = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

        private bool m_should_close = false;
        private NotifyIcon m_trayIcon;
        private System.Windows.Forms.ContextMenu m_trayMenu;
        public string m_user_name;
        public int m_tot_cnt;
        public int m_done_cnt;
        public string m_last_file = "";

        public string m_prev_A = "";
        public string m_prev_B = "";
        public string m_prev_C = "";

        object[,]   m_records;
        public xlsf m_xls = new xlsf();
        public List<Order> m_orders = new List<Order>();
        public List<Keys> m_key_hist = new List<Keys>();
        public Dictionary<string, int> m_order_tot = new Dictionary<string, int>();
        public Dictionary<string, int> m_order_done = new Dictionary<string, int>();
        public Dictionary<string, int> m_art_tot = new Dictionary<string, int>();
        public Dictionary<string, int> m_art_done = new Dictionary<string, int>();
        public Dictionary<string, bool> m_back_order = new Dictionary<string, bool>();
        System.Windows.Forms.Timer m_timer = new System.Windows.Forms.Timer();
        int m_cnt = 0;
        public MainWindow()
        {
            if (App.is_closing)
                return;

            InitializeComponent();
            //ShowInTaskbar = false; // Remove from taskbar.

            SetStartup();
            m_user_name = get_current_user();

            // create a tray menu
            m_trayMenu = new System.Windows.Forms.ContextMenu();
            m_trayIcon = new NotifyIcon();
            m_trayIcon.Text = "Order Packing Helper";
            try
            {
                m_trayIcon.Icon = new Icon("icon.ico");
            }
            catch (Exception e)
            {
                m_trayIcon.Icon = System.Drawing.SystemIcons.WinLogo;
            }
            m_trayIcon.ContextMenu = m_trayMenu;
            m_trayIcon.Visible = true;
            m_trayIcon.DoubleClick += OnMenuPacking;
            m_trayMenu.MenuItems.Add("Start Packing", OnMenuPacking);
            m_trayMenu.MenuItems.Add("Exit", OnExit);
            

            // initialize hooks
            init_hooks();

            // greetings
            string greet = "";
            if (DateTime.Now.Hour < 11)
                greet = "Good morning, " + m_user_name;
            else if (DateTime.Now.Hour < 18)
                greet = "Good afternoon, " + m_user_name;
            else
                greet = "Good evening, " + m_user_name;
            show_info(greet, "Order Packing Helper Started");

            m_timer.Interval = 5000;
            m_timer.Tick += M_timer_Tick;
            m_timer.Start();
        }

        private void M_timer_Tick(object sender, EventArgs e)
        {
            m_cnt = (m_cnt + 1) % 2;
            
            string backup = Directory.GetCurrentDirectory() + $"\\backup_{m_cnt}.xls";
            try
            {
                save_table(backup);
            }
            catch (Exception ex)
            {

            }
        }

        public void init_hooks()
        {
            logger.Info(">> Init hooks");
            UserActivityMonitor.HookManager.KeyDown += hook_activity;
            logger.Info("<< Init hooks");
        }

        private void hook_activity(object sender, System.Windows.Forms.KeyEventArgs e)
        {
            if (Visibility == Visibility.Hidden)
                return;

            Keys key = e.KeyCode;
            m_key_hist.Add(key);
            if (m_key_hist.Count > 16)
                m_key_hist.RemoveAt(0);
            if (m_key_hist.Count < 16)
                return;

            if(key == Keys.Enter || key == Keys.Return)
            {
                if (m_key_hist[0] == Keys.A || m_key_hist[0] == Keys.B || m_key_hist[0] == Keys.C)
                {
                    int i;
                    for (i = 1; i <= 14; i++)
                    {
                        if (is_digit_key(m_key_hist[i]) == false)
                            break;
                    }
                    if (i == 15)
                    {
                        string key_str = "";
                        for (int j = 1; j < 9; j++)
                        {
                            key_str += numkey2str(m_key_hist[j]);
                        }

                        do_process(m_key_hist[0], key_str);
                    }
                }
            }
            return;
        }
        
        public string add_dots(string article)
        {

            string str = article.Substring(0, 3) + "." + article.Substring(3, 3) + "." + article.Substring(6);
            return str;
        }

        public void play_beep()
        {
            var notificationSound = new SoundPlayer(Properties.Resources.beep);
            notificationSound.Play();
        }
        public void do_process(Keys p_scanner, string p_article)
        {

            bool found = false;
            for(int i = 0; i < m_orders.Count; i ++)
            {
                Order ord = m_orders[i];

                if (is_equal(ord.m_article, p_article))
                {
                    if(ord.m_delivered < ord.m_quantity)
                    {

                        found = true;

                        play_beep();

                        string art_str = ord.m_article + ord.m_order_num;

                        ord.m_delivered++;
                        m_order_done[ord.m_order_num]++;
                        m_art_done[art_str]++;
                        m_done_cnt++;

                        

                        refresh_whole_prog();

                        if(p_scanner == Keys.A)
                        {
                            ui_art_num_A.Content = add_dots(p_article);
                            ui_done_order_A.Text = m_order_done[ord.m_order_num].ToString();
                            ui_ordersize_A.Text = m_order_tot[ord.m_order_num].ToString();
                            ui_done_art_A.Text = m_art_done[art_str].ToString();
                            ui_artsize_A.Text = m_art_tot[art_str].ToString();

                            ui_ord_num_A.Content = ord.m_order_num;

                            bool val = false;
                            if (m_back_order.TryGetValue(ord.m_order_num, out val) == true)
                                ui_ord_num_A.Background = new SolidColorBrush(System.Windows.Media.Color.FromArgb(0x66, 0xE6, 0x26, 0x26));
                            else
                                ui_ord_num_A.ClearValue(System.Windows.Controls.Control.BackgroundProperty);

                            if (m_order_done[ord.m_order_num] == 1)
                                ui_state_first_A.Visibility = Visibility.Visible;
                            else
                                ui_state_first_A.Visibility = Visibility.Hidden;

                            if (m_order_done[ord.m_order_num] == m_order_tot[ord.m_order_num])
                                ui_state_last_A.Visibility = Visibility.Visible;
                            else
                                ui_state_last_A.Visibility = Visibility.Hidden;

                            ui_last_scanned_A.Content = m_prev_A;
                            m_prev_A = String.Format("Previous: {0} of {1}", ui_art_num_A.Content, ui_ord_num_A.Content);
                        }
                        else if (p_scanner == Keys.B)
                        {
                            ui_art_num_B.Content = add_dots(p_article);
                            ui_done_order_B.Text = m_order_done[ord.m_order_num].ToString();
                            ui_ordersize_B.Text = m_order_tot[ord.m_order_num].ToString();
                            ui_done_art_B.Text = m_art_done[art_str].ToString();
                            ui_artsize_B.Text = m_art_tot[art_str].ToString();
                            ui_ord_num_B.Content = ord.m_order_num;

                            bool val = false;
                            if (m_back_order.TryGetValue(ord.m_order_num, out val) == true)
                                ui_ord_num_B.Background = new SolidColorBrush(System.Windows.Media.Color.FromArgb(0x66, 0xE6, 0x26, 0x26));
                            else
                                ui_ord_num_B.ClearValue(System.Windows.Controls.Control.BackgroundProperty);

                            if (m_order_done[ord.m_order_num] == 1)
                                ui_state_first_B.Visibility = Visibility.Visible;
                            else
                                ui_state_first_B.Visibility = Visibility.Hidden;

                            if (m_order_done[ord.m_order_num] == m_order_tot[ord.m_order_num])
                                ui_state_last_B.Visibility = Visibility.Visible;
                            else
                                ui_state_last_B.Visibility = Visibility.Hidden;

                            ui_last_scanned_B.Content = m_prev_B;
                            m_prev_B = String.Format("Previous: {0} of {1}", ui_art_num_B.Content, ui_ord_num_B.Content);
                        }
                        else if (p_scanner == Keys.C)
                        {
                            ui_art_num_C.Content = add_dots(p_article);
                            ui_done_order_C.Text = m_order_done[ord.m_order_num].ToString();
                            ui_ordersize_C.Text = m_order_tot[ord.m_order_num].ToString();
                            ui_done_art_C.Text = m_art_done[art_str].ToString();
                            ui_artsize_C.Text = m_art_tot[art_str].ToString();
                            ui_ord_num_C.Content = ord.m_order_num;

                            bool val = false;
                            if (m_back_order.TryGetValue(ord.m_order_num, out val) == true)
                                ui_ord_num_C.Background = new SolidColorBrush(System.Windows.Media.Color.FromArgb(0x66, 0xE6, 0x26, 0x26));
                            else
                                ui_ord_num_C.ClearValue(System.Windows.Controls.Control.BackgroundProperty);

                            if (m_order_done[ord.m_order_num] == 1)
                                ui_state_first_C.Visibility = Visibility.Visible;
                            else
                                ui_state_first_C.Visibility = Visibility.Hidden;

                            if (m_order_done[ord.m_order_num] == m_order_tot[ord.m_order_num])
                                ui_state_last_C.Visibility = Visibility.Visible;
                            else
                                ui_state_last_C.Visibility = Visibility.Hidden;

                            ui_last_scanned_C.Content = m_prev_C;
                            m_prev_C = String.Format("Previous: {0} of {1}", ui_art_num_C.Content, ui_ord_num_C.Content);
                        }

                        return;
                    }
                }
            }

            if(found == false)
            {
                string init = "";
                if (p_scanner == Keys.A)
                {
                    ui_art_num_A.Content = add_dots(p_article);
                    ui_done_order_A.Text = init; ui_ordersize_A.Text = init;
                    ui_done_art_A.Text = init; ui_artsize_A.Text = init;
                    ui_ord_num_A.Content = "XXXX";

                    ui_state_first_A.Visibility = Visibility.Hidden;
                    ui_state_last_A.Visibility = Visibility.Hidden;
                    ui_last_scanned_A.Content = m_prev_A;
                }
                else if (p_scanner == Keys.B)
                {
                    ui_art_num_B.Content = add_dots(p_article);
                    ui_done_order_B.Text = init; ui_ordersize_B.Text = init;
                    ui_done_art_B.Text = init; ui_artsize_B.Text = init;
                    ui_ord_num_B.Content = "XXXX";

                    ui_state_first_B.Visibility = Visibility.Hidden;
                    ui_state_last_B.Visibility = Visibility.Hidden;
                    ui_last_scanned_B.Content = m_prev_B;
                }
                else if (p_scanner == Keys.C)
                {
                    ui_art_num_C.Content = add_dots(p_article);
                    ui_done_order_C.Text = init; ui_ordersize_C.Text = init;
                    ui_done_art_C.Text = init; ui_artsize_C.Text = init;
                    ui_ord_num_C.Content = "XXXX";

                    ui_state_first_C.Visibility = Visibility.Hidden;
                    ui_state_last_C.Visibility = Visibility.Hidden;
                    ui_last_scanned_C.Content = m_prev_C;
                }

                show_info("Order not found", p_article);
            }
        }

        private bool is_equal(string code1, string code2)
        {
            int len = Math.Min(code1.Length, code2.Length);
            for(int i = 0; i < len; i ++)
            {
                if (code1[code1.Length - i - 1] != code2[code2.Length - i - 1])
                    return false;
            }
            return true;
        }
        private bool is_digit_key(Keys key)
        {
            if (key >= Keys.D0 && key <= Keys.D9 ||
               key >= Keys.NumPad0 && key <= Keys.NumPad9)
                return true;
            return false;
        }

        private string numkey2str(Keys key)
        {
            if (key >= Keys.D0 && key <= Keys.D9)
                return (key - Keys.D0).ToString();
            if (key >= Keys.NumPad0 && key <= Keys.NumPad9)
                return (key - Keys.NumPad0).ToString();
            return "";
        }
        private void OnExit(object sender, EventArgs e)
        {
            m_should_close = true;
            show_info("Thank you", "Contact us : pck2016217@gmail.com");
            Close();
        }
        private void OnMenuPacking(object sender, EventArgs e)
        {
            if (Visibility == Visibility.Visible)
            {
                Visibility = Visibility.Hidden;
                m_trayMenu.MenuItems[0].Text = "Start Packing";
                m_timer.Stop();
            }
            else
            {
                Visibility = Visibility.Visible;

                init_controls();
                if(m_last_file != "")
                {
                    m_xls.OpenFile(m_last_file);
                    ui_order_path.Content = m_last_file;
                    prepare_work();
                }

                m_timer.Start();
                m_trayMenu.MenuItems[0].Text = "Stop Packing";
            }
        }

        private void SetStartup()
        {
            RegistryKey rk = Registry.CurrentUser.OpenSubKey
                ("SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Run", true);
            rk.SetValue("Order Packaging Helper", System.Windows.Forms.Application.ExecutablePath);
        }

        private void Window_Loaded(object sender, RoutedEventArgs e)
        {
            init_controls();
        }

        public void show_info(string title, string text)
        {
            m_trayIcon.BalloonTipIcon = ToolTipIcon.Info;
            m_trayIcon.BalloonTipTitle = title;
            m_trayIcon.BalloonTipText = text;
            m_trayIcon.ShowBalloonTip(3000);
        }

        public string get_current_user()
        {
            string identity = System.Security.Principal.WindowsIdentity.GetCurrent().Name;
            logger.Info(" >< get username " + identity);
            return identity.Substring(identity.LastIndexOf('\\') + 1);
        }

        private void Image_MouseDown_Exit(object sender, MouseButtonEventArgs e)
        {
            try
            {
                OnMenuPacking(null, null);
                // save to tmp file
                string tmp_path = m_last_file.Substring(0, m_last_file.LastIndexOf('\\') + 1) + "temp.xlsx";
                save_table(tmp_path);
                m_xls.CloseFile(true);
                m_orders.Clear();
            }
            catch(Exception ex)
            {
                System.Diagnostics.Debug.WriteLine(ex.Message + "\n" + ex.StackTrace);
            }
        }

        private void init_controls()
        {
            string init = "";
            ui_art_num_A.Content = ui_art_num_B.Content = ui_art_num_C.Content = "";
            ui_done_order_A.Text = ui_done_order_B.Text = ui_done_order_C.Text = init;
            ui_ordersize_A.Text = ui_ordersize_B.Text = ui_ordersize_C.Text = init;
            ui_done_art_A.Text = ui_done_art_B.Text = ui_done_art_C.Text = init;
            ui_artsize_A.Text = ui_artsize_B.Text = ui_artsize_C.Text = init;

            ui_tot_process.Content = "";
            ui_ord_num_A.Content = ui_ord_num_B.Content = ui_ord_num_C.Content = "";
            ui_ord_num_A.ClearValue(System.Windows.Controls.Control.BackgroundProperty);
            ui_ord_num_B.ClearValue(System.Windows.Controls.Control.BackgroundProperty);
            ui_ord_num_C.ClearValue(System.Windows.Controls.Control.BackgroundProperty);
            ui_state_first_A.Visibility = ui_state_first_B.Visibility = ui_state_first_C.Visibility = Visibility.Hidden;
            ui_state_last_A.Visibility = ui_state_last_B.Visibility = ui_state_last_C.Visibility = Visibility.Hidden;
            ui_order_path.Content = "Order table is not selected yet";
            ui_bar_prog.Value = 0;
        }
        private void Image_MouseDown_Browse(object sender, MouseButtonEventArgs e)
        {
            Microsoft.Win32.FileDialog dialog = new Microsoft.Win32.OpenFileDialog();
            dialog.DefaultExt = ".xlsx"; // Default file extension 
            dialog.Filter = "Excel Worksheets|*.xls;*.xlsx"; // Filter files by extension
            if (dialog.ShowDialog() == true)
            {
                var fullPath = dialog.FileName;
                m_xls.OpenFile(fullPath.ToString());

                init_controls();
                ui_order_path.Content = fullPath.ToString();
                m_last_file = fullPath.ToString();
                prepare_work();
            }
        }

        public void prepare_work()
        {
            load_table();
            refresh_whole_prog();
        }

        public void save_table(string filename)
        {
            try
            {
                if (m_xls.excelApp.ActiveWorkbook != null && m_orders.Count > 0)
                {
                    for (int i = 0; i < m_orders.Count; i++)
                    {
                        Order ord = m_orders[i];
                        if (ord.m_delivered >= ord.m_quantity && ord.m_delivered > 0)
                            m_records[ord.m_row, 6] = "X";
                        else if (ord.m_delivered > 0)
                            m_records[ord.m_row, 6] = ord.m_delivered.ToString();
                    }

                    Microsoft.Office.Interop.Excel.Worksheet gXlWs = (Microsoft.Office.Interop.Excel.Worksheet)m_xls.excelApp.ActiveWorkbook.Worksheets.get_Item(1);
                    Microsoft.Office.Interop.Excel.Range range = gXlWs.UsedRange;// get_Range("A1", "F188000");
                    range.Value2 = m_records;
                    m_xls.Save();

                    string org_fname = m_xls.excelApp.ActiveWorkbook.FullName;
                    System.IO.File.Copy(org_fname, filename, true);
                }
            }
            catch (Exception ex)
            {

            }
        }
        public void load_table()
        {
            m_orders.Clear();
            m_order_tot.Clear(); m_order_done.Clear();
            m_art_tot.Clear(); m_art_done.Clear();
            m_back_order.Clear();
            ui_load_prog.Value = 0;
            ui_load_prog.Visibility = Visibility.Visible;

            int tot = (int)m_xls.GetVrticlTotalCell();
            Microsoft.Office.Interop.Excel.Worksheet gXlWs = (Microsoft.Office.Interop.Excel.Worksheet)m_xls.excelApp.ActiveWorkbook.Worksheets.get_Item(1);
            Microsoft.Office.Interop.Excel.Range range = gXlWs.UsedRange;// get_Range("A1", "F188000");

            m_records = (object[,])range.Value2;
            m_tot_cnt = m_done_cnt = 0;
            for(int row = 2; row < tot; row ++)
            {
                string back = obj2str(m_records[row, 5]);
                if (back.Trim() == "backorder")
                {
                    m_back_order[obj2str(m_records[row, 7])] = true;
                    continue;
                }
                Order ord = new Order();               
                ord.m_row = row;
                ord.m_article = obj2str(m_records[row, 1]);
                ord.m_quantity = obj2int(m_records[row, 3]);
                m_tot_cnt += ord.m_quantity;
                ord.m_delivered = obj2int(m_records[row, 6]);
                if(obj2str(m_records[row, 6]) == "X")
                    ord.m_delivered  = ord.m_quantity;
                m_done_cnt += ord.m_delivered;

                ord.m_order_num = obj2str(m_records[row, 7]);
                m_orders.Add(ord);
                ui_load_prog.Dispatcher.Invoke(() => ui_load_prog.Value = 100 * (row - 1.0) / tot, DispatcherPriority.Background);

                int val;
                if (m_order_tot.TryGetValue(ord.m_order_num, out val) == false)
                    m_order_tot[ord.m_order_num] = ord.m_quantity;
                else
                    m_order_tot[ord.m_order_num] += ord.m_quantity;

                if (m_order_done.TryGetValue(ord.m_order_num, out val) == false)
                    m_order_done[ord.m_order_num] = ord.m_delivered;
                else
                    m_order_done[ord.m_order_num] += ord.m_delivered;

                string art_str = ord.m_article + ord.m_order_num;
                if (m_art_tot.TryGetValue(art_str, out val) == false)
                    m_art_tot[art_str] = ord.m_quantity;
                else
                    m_art_tot[art_str] += ord.m_quantity;

                if (m_art_done.TryGetValue(art_str, out val) == false)
                    m_art_done[art_str] = ord.m_delivered;
                else
                    m_art_done[art_str] += ord.m_delivered;
            }

            ui_load_prog.Visibility = Visibility.Hidden;
        }
        public void refresh_whole_prog()
        {
            ui_tot_process.Content = String.Format("{0}/{1}", m_done_cnt, m_tot_cnt);
            ui_bar_prog.Value = m_done_cnt * 100 / m_tot_cnt;
        }

        private int obj2int(object x)
        {
            if (x == null)
                return 0;
            else if (x.ToString() == "X")
                return 0;
            else
            {
                try
                {
                    return Int32.Parse(x.ToString());
                }
                catch (Exception e)
                {
                    return 0;
                }
            }   
        }

        private string obj2str(object x)
        {
            if (x == null)
                return "";
            else
                return x.ToString();
        }

        private void Window_LostFocus(object sender, RoutedEventArgs e)
        {
            //save_table();
        }

        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            if(m_should_close == false)
            {
                e.Cancel = true;
                show_info("You can close using the tray menu in the taskbar.", "Order Pack Helper");
            }
            save_table(m_last_file);
        }
    }

    public class Order
    {
        public int m_row;
        public string m_article;
        public int m_quantity;
        public int m_delivered;
        public string m_order_num;
    }
}
