using Microsoft.EntityFrameworkCore;
using System.Collections.ObjectModel;
using System.Security.Cryptography.Pkcs;
using System.Windows;
using System.Collections.Generic;
using System.Windows.Markup;
using Newtonsoft.Json;
using System.IO;

namespace SQLiteApp
{
    public partial class MainWindow : Window
    {
        ApplicationContext db = new ApplicationContext();
        public MainWindow()
        {
            InitializeComponent();

            Loaded += MainWindow_Loaded;
        }

        // при загрузке окна
        private void MainWindow_Loaded(object sender, RoutedEventArgs e)
        {
            // гарантируем, что база данных создана
            db.Database.EnsureCreated();
            // загружаем данные из БД
            db.Users.Load();
            // и устанавливаем данные в качестве контекста
            DataContext = db.Users.Local.ToObservableCollection();
        }

        // добавление
        private void Add_Click(object sender, RoutedEventArgs e)
        {
            UserWindow UserWindow = new UserWindow(new User());
            if (UserWindow.ShowDialog() == true)
            {
                User User = UserWindow.User;
                db.Users.Add(User);
                db.SaveChanges();
            }
        }
        // редактирование
        private void Edit_Click(object sender, RoutedEventArgs e)
        {
            // получаем выделенный объект
            User? user = usersList.SelectedItem as User;
            // если ни одного объекта не выделено, выходим
            if (user is null) return;

            UserWindow UserWindow = new UserWindow(new User
            {
                Id = user.Id,
                Name = user.Name,
                Family = user.Family,
                Patronymicl = user.Patronymicl,
                Telephone = user.Telephone,
                Department = user.Department,
                Birthday = user.Birthday,
            });

            if (UserWindow.ShowDialog() == true)
            {
                // получаем измененный объект
                user = db.Users.Find(UserWindow.User.Id);
                if (user != null)
                {
                    user.Id = UserWindow.User.Id;
                    user.Name = UserWindow.User.Name;
                    user.Family = UserWindow.User.Family;
                    user.Patronymicl = UserWindow.User.Patronymicl;
                    user.Telephone = UserWindow.User.Telephone;
                    user.Birthday = UserWindow.User.Birthday;
                    user.Department = UserWindow.User.Department;
                    db.SaveChanges();
                    usersList.Items.Refresh();
                }
            }
        }
        // удаление
        private void Delete_Click(object sender, RoutedEventArgs e)
        {
            // получаем выделенный объект
            User? user = usersList.SelectedItem as User;
            // если ни одного объекта не выделено, выходим
            if (user is null) return;
            db.Users.Remove(user);
            db.SaveChanges();
        }
        private void CreateExcelfile(object sender, RoutedEventArgs e)
        {
            EEExlEEE report = new EEExlEEE();
            report.CreateExcelfile(db.Users.Local.ToObservableCollection());
        }
        private void CreateJsonfile(object sender, RoutedEventArgs e)
        {
            var Json = JsonConvert.SerializeObject(DataContext);
            string DRpath = "Reports";
            if (Directory.Exists(DRpath) == false)
            {
                Directory.CreateDirectory(DRpath);
            }
            string exportfile = "Report.json";
            DRpath = System.IO.Path.Combine(DRpath, exportfile);
            if (File.Exists(DRpath))
                File.Delete(DRpath);
            FileStream objFileStrm = File.Create(DRpath);
            objFileStrm.Close();
            File.WriteAllText(DRpath, Json.ToString());
        }
    }
}