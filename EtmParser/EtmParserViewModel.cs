using Microsoft.Win32;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Configuration;
using System.Globalization;
using System.Linq;
using System.Net.Http;
using System.Reflection;
using System.Runtime.CompilerServices;
using System.Security.Policy;
using System.Text;
using System.Text.Json;
using System.Text.Json.Nodes;
using System.Text.Json.Serialization;
using System.Threading.Tasks;
using System.Windows.Data;
using ClosedXML.Excel;
using System.Windows.Threading;
using DocumentFormat.OpenXml.Spreadsheet;
using System.Windows;

namespace EtmParser
{
    internal class EtmParserViewModel : INotifyPropertyChanged
    {
        private Configuration _configFile;
        private ApiRequester _apiRequester;

        private int _totalGoods;
        private int _parsedGoods;

        public event PropertyChangedEventHandler? PropertyChanged;

        public string Login { get; set; }
        public string Password { get; set; }
        public CodesType SelectedCodesType { get; set; }
        public string Status
        {
            get
            {
                return _apiRequester.Status switch
                {
                    ApiRequesterStatus.Ready => "Готов к работе",
                    ApiRequesterStatus.NeedLogin => "Авторизация не пройдена",
                    ApiRequesterStatus.Working => "В работе",
                    _ => "Не готов",
                };
            }
        }


        private string _fileName;
        public string FileName 
        { 
            get => _fileName;
            set 
            { 
                _fileName = value;
                OnPropertyChanged();
            } 
        }

        public string Progress => $"{_parsedGoods}/{_totalGoods}";

        private bool _goodsInfoRequest = true;
        public bool GoodsInfoRequest
        {
            get => _goodsInfoRequest;
            set
            {
                _goodsInfoRequest = value;
                OnPropertyChanged();
            }
        }

        private bool _remainInfoRequest = true;
        public bool RemainInfoRequest
        {
            get => _remainInfoRequest;
            set
            {
                _remainInfoRequest = value;
                OnPropertyChanged();
            }
        }

        public EtmParserViewModel()
        {
            Initialization();
        }

        private void Initialization()
        {
            _configFile = ConfigurationManager.OpenExeConfiguration(ConfigurationUserLevel.None);

            var server = _configFile.AppSettings.Settings["Server"].Value;
            _apiRequester = new ApiRequester(server);

            Login = _configFile.AppSettings.Settings["Login"].Value;
            Password = _configFile.AppSettings.Settings["Password"].Value;

            if (!string.IsNullOrEmpty(_configFile.AppSettings.Settings["CodeType"].Value))
                SelectedCodesType = Enum.Parse<CodesType>(_configFile.AppSettings.Settings["CodeType"].Value);

            var tokenGeneratedDateTime = _configFile.AppSettings.Settings["TokenGeneratedTime"].Value;
            var token = _configFile.AppSettings.Settings["ApiToken"].Value;

            if (!string.IsNullOrEmpty(tokenGeneratedDateTime) && !string.IsNullOrEmpty(token))
            {
                var time = DateTime.Parse(tokenGeneratedDateTime);
                var diff = DateTime.Now - time;
                if (diff.TotalSeconds < 600)
                {
                    _apiRequester.SetToken(token);
                    OnPropertyChanged(nameof(Status));
                    return;
                }
            }

            if (!string.IsNullOrEmpty(Login) && !string.IsNullOrEmpty(Password))
                Authorize();
        }

        public void UpdateSettings()
        {
            var settings = _configFile.AppSettings.Settings;

            if (settings["Login"].Value != Login || settings["Password"].Value != Password)
                Authorize();

            settings["Login"].Value = Login;
            settings["Password"].Value = Password;
            settings["CodeType"].Value = SelectedCodesType.ToString();

            _configFile.Save(ConfigurationSaveMode.Modified);
            ConfigurationManager.RefreshSection(_configFile.AppSettings.SectionInformation.Name);
        }

        //Получение токена. Вызывается при запуске программы или после заполнеия логина и пароля
        private void Authorize()
        {
            var tokenGeneratedDateTime = _configFile.AppSettings.Settings["TokenGeneratedTime"].Value;

            //Проверяем время, получаем свежий токен по возможности
            if (!string.IsNullOrEmpty(tokenGeneratedDateTime))
            {
                var time = DateTime.Parse(tokenGeneratedDateTime);
                var diff = DateTime.Now - time;

                if (diff.TotalSeconds < 130)
                {
                    return;
                }
            }

            var token = _apiRequester.AuthorizationRequest(Login, Password);
            if (string.IsNullOrEmpty(token))
            {
                OnPropertyChanged(nameof(Status));
                return;
            }

            //Сохранение токена в конфиге
            var settings = _configFile.AppSettings.Settings;
            settings["ApiToken"].Value = token;
            settings["TokenGeneratedTime"].Value = DateTime.Now.ToString();

            _configFile.Save(ConfigurationSaveMode.Modified);
            ConfigurationManager.RefreshSection(_configFile.AppSettings.SectionInformation.Name);

            OnPropertyChanged(nameof(Status));
        }

        public async void ParseGoods()
        {
            //Выбор файла
            var openFileDialog = new OpenFileDialog();
            openFileDialog.Filter = "excel files (*.xlsx)|*.xlsx";
            openFileDialog.RestoreDirectory = true;

            if (openFileDialog.ShowDialog() != true)
                return;

            _apiRequester.Status = ApiRequesterStatus.Working;
            OnPropertyChanged(nameof(Status));
            await ProcessExcelAsync(openFileDialog.FileName);
        }

        //Метод для заполения Excel таблицы
        private async Task ProcessExcelAsync(string filePath)
        {
            await Task.Run(() =>
            {
                using var workbook = new XLWorkbook(filePath);
                var worksheet = workbook.Worksheet(1);

                _parsedGoods = 0;
                _totalGoods = GetTotalGoodsCount(worksheet);

                int row = 2;
                var goodNotFoundList = new List<string>();
                var storeNotFound = new List<string>();

                var saveCount = 0;

                while (!worksheet.Cell(row, 1).IsEmpty())
                {
                    var inputValue = worksheet.Cell(row, 1).GetString();
                    var mnf = worksheet.Cell(row, 2).GetString();

                    if (GoodsInfoRequest)
                    {
                        var goodsInfo = _apiRequester.GoodsRequest(inputValue, SelectedCodesType.ToString(), mnf);
                        if (goodsInfo != null && goodsInfo.Data != null)
                        {
                            if (goodsInfo.Data.PackData != null)
                            {
                                worksheet.Cell(row, 4).Value = goodsInfo.Data.PackData[0].Weight ?? string.Empty;
                                worksheet.Cell(row, 5).Value = goodsInfo.Data.PackData[0].Length ?? string.Empty;
                                worksheet.Cell(row, 6).Value = goodsInfo.Data.PackData[0].Width ?? string.Empty;
                                worksheet.Cell(row, 7).Value = goodsInfo.Data.PackData[0].Height ?? string.Empty;
                            }

                            worksheet.Cell(row, 8).Value = goodsInfo.Data.Code ?? string.Empty;
                            worksheet.Cell(row, 9).Value = goodsInfo.Data.GoodName ?? string.Empty;
                        }
                    }

                    if (RemainInfoRequest)
                    {
                        //Запрос по остаткам товара.
                        var remainInfo = _apiRequester.RemainRequest(inputValue, SelectedCodesType.ToString(), mnf);

                        //Заполенение Excel таблицы
                        if(remainInfo != null && remainInfo.Data != null && remainInfo.Data.InfoStores != null)
                        {
                            var infoStore = remainInfo.Data.InfoStores.FirstOrDefault(s => s.StoreCode == 12);
                            if (infoStore != null) 
                            {
                                worksheet.Cell(row, 10).Value = infoStore.Remain;
                            }
                            else
                            {
                                worksheet.Cell(row, 10).Value = "Склад не найден";
                                worksheet.Cell(row, 10).Style.Fill.BackgroundColor = XLColor.Orange;
                                storeNotFound.Add(inputValue);
                            }
                        }
                        else
                        {
                            worksheet.Cell(row, 10).Value = "Товар не найден";
                            worksheet.Cell(row, 10).Style.Fill.BackgroundColor = XLColor.Red;
                            goodNotFoundList.Add(inputValue);
                        }
                    }

                    row++;
                    _parsedGoods++;

                    Application.Current.Dispatcher.Invoke(() =>
                    {
                        OnPropertyChanged(nameof(Progress));
                    });

                    //Сохранение файла через каждые 100 запросов
                    saveCount++;
                    if(saveCount == 100)
                    {
                        saveCount = 0;
                        workbook.Save();
                    }
                }

                if (goodNotFoundList.Any())
                {
                    row++;
                    worksheet.Cell(row, 1).Value = "Не найдены товары:";
                    row++;
                    foreach (var good in goodNotFoundList) 
                    {
                        worksheet.Cell(row, 1).Value = good;
                        row++;
                    }
                }

                if (storeNotFound.Any())
                {
                    row++;
                    worksheet.Cell(row, 1).Value = "Не найдены склад Урал для товаров:";
                    row++;
                    foreach (var store in storeNotFound)
                    {
                        worksheet.Cell(row, 1).Value = store;
                        row++;
                    }
                }

                _apiRequester.Status = ApiRequesterStatus.Ready;
                Application.Current.Dispatcher.Invoke(() =>
                {
                    OnPropertyChanged(nameof(Status));
                });
                workbook.Save();
            });
        }

        static int GetTotalGoodsCount(IXLWorksheet worksheet)
        {
            int row = 2;
            int count = 0;

            while (true)
            {
                var cell = worksheet.Cell(row, 1);
                if (cell.IsEmpty())
                    break;

                count++;
                row++;
            }

            return count;
        }

        protected void OnPropertyChanged([CallerMemberName] string name = null)
        {
            PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
        }
    }


    //Класс реализующий отправку запросов к API
    internal class ApiRequester
    {
        HttpClient _client = new HttpClient();
        private string _server;

        public string Token { get; set; }
        public ApiRequesterStatus Status { get; set; } = ApiRequesterStatus.NeedLogin; //Статус для обновления UI

        public ApiRequester(string server)
        {
            _server = server;
        }

        public void SetToken(string token)
        {
            Token = token;
            Status = ApiRequesterStatus.Ready;
        }

        //Запрос авторизации
        public string? AuthorizationRequest(string login, string password)
        {
            Token = string.Empty;
            HttpResponseMessage response = _client.PostAsync($"{_server}/user/login?log={login}&pwd={password}", null).Result;

            if (response.IsSuccessStatusCode)
            {
                var jsonData = JsonSerializer.Deserialize<AuthorizationResponse>(response.Content.ReadAsStringAsync().Result);
                var token = jsonData?.Data.Session;
                if (!string.IsNullOrEmpty(token))
                {
                    Token = token;
                    Status = ApiRequesterStatus.Ready;
                }
                else
                {
                    Status = ApiRequesterStatus.NeedLogin;
                }
                return token;
            }

            Status = ApiRequesterStatus.NeedLogin;
            return null;
        }

        //Запрос по характеристикам товара
        public Good? GoodsRequest(string id, string type, string? mnf = null)
        {
            var request = $"{_server}/goods/{id}?type={type}&session-id={Token}";

            if (!string.IsNullOrEmpty(mnf))
                request = request + $"&mnf={mnf}";

            HttpResponseMessage response = _client.GetAsync(request).Result;
            Thread.Sleep(1200);

            if (response.IsSuccessStatusCode)
            {
                return JsonSerializer.Deserialize<Good>(response.Content.ReadAsStringAsync().Result);
            }

            return null;
        }

        //Запрос по остатку товара
        public Remain? RemainRequest(string id, string type, string? mnf = null)
        {
            var request = $"{_server}/goods/{id}/remains?type={type}&session-id={Token}";

            if (!string.IsNullOrEmpty(mnf))
                request = request + $"&mnf={mnf}"; // Код производителя, заполняется для type=mnf

            //Запрос с типом кода ETM выглядит так: https://ipro.etm.ru/api/v1/goods/1219732/remains?type=Etm&session-id=XXXX

            HttpResponseMessage response = _client.GetAsync(request).Result; //Запрос к API
            Thread.Sleep(1200);

            if (response.IsSuccessStatusCode)
            {
                var r = response.Content.ReadAsStringAsync().Result; //Чтение ответа
                return JsonSerializer.Deserialize<Remain>(r); //Десериализация JSON в специально определенный класс
            }

            return null;
        }
    }

    //Ниже конвертеры и енамы для UI части

    public class CodesTypeEnumToDescriptionValueConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo info)
        {
            var type = typeof(CodesType);
            var name = Enum.GetName(type, value);
            FieldInfo fi = type.GetField(name);
            var descriptionAttrib = (DescriptionAttribute)Attribute.GetCustomAttribute(fi, typeof(DescriptionAttribute));

            return descriptionAttrib.Description;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotSupportedException();
        }
    }

    public class StatusToAvailableSaveSettingsConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo info)
        {
            if (value is not string status)
                return false;

            if(status == "Готов к работе" || status == "Авторизация не пройдена")
                return true;

            return false;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotSupportedException();
        }
    }

    public class StatusToAvailableParsingConverter : IValueConverter
    {
        public object Convert(object value, Type targetType, object parameter, CultureInfo info)
        {
            if (value is not string status)
                return false;

            return status == "Готов к работе" ? true : false;
        }

        public object ConvertBack(object value, Type targetType, object parameter, CultureInfo culture)
        {
            throw new NotSupportedException();
        }
    }

    public enum CodesType
    {
        [Description("Коды клиента")]
        Cli,

        [Description("Коды ЭТМ")]
        Etm,

        [Description("Коды производителя")]
        mnf
    }

    public enum ApiRequesterStatus
    {
        [Description("Необходима автоизация")]
        NeedLogin,

        [Description("Готов к работе")]
        Ready,

        [Description("Работает")]
        Working
    }
}
