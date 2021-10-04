using NLog;
using System;
using System.IO;
using Topshelf;
using Npoi.Mapper;
using System.Collections.Generic;
using System.Linq;
using WinService.EqpService;
using System.ServiceModel;
using System.Configuration;
using Import_Eqp;

namespace WinService
{
    internal interface IService
    {
        void Start();
        void Stop();
    }
    internal class Service : IService
    {
        public void Start()
        {

        }
        public void Stop()
        {

        }
    }
    internal class ServiceConfiguration
    {
        public static void Configure()
        {
            HostFactory.Run(h =>
            {
                h.Service<Service>(s =>
                {
                    s.ConstructUsing(name => new Service());
                    s.WhenStarted(c => c.Start());
                    s.WhenStopped(c => c.Stop());
                });
                h.StartAutomatically();
                h.RunAsLocalSystem();
                h.SetServiceName("Greenatom.Lims.Eqp_Import");
                h.SetDisplayName("Greenatom.Lims.Eqp_Import");
                h.SetDescription("Служба автоматического импорта из УРПО в ЛИМС ");
                h.EnableServiceRecovery(recoveryOption =>
                {
                    recoveryOption.RestartService(0);
                });
                h.UseNLog();
            });
        }
        public class Program
        {
            public static void Main()
            {
                FileSystemWatcher();
                Configure();
            }
            /// <summary> для отслеживания изменений в указанном каталоге </summary>
            public static void FileSystemWatcher()
            {
                string path = ConfigurationManager.AppSettings["path"];
                string filter = ConfigurationManager.AppSettings["filter"];
                FileSystemWatcher fsw = new(path, filter)
                {
                    IncludeSubdirectories = true,
                    EnableRaisingEvents = true,
                    NotifyFilter = NotifyFilters.LastAccess | NotifyFilters.LastWrite
                                   | NotifyFilters.FileName | NotifyFilters.DirectoryName
                };
                fsw.Changed += OnChanged;
                fsw.EnableRaisingEvents = true;
            }
            private static IList<Equipment> _equipments;
            public static Logger log = LogManager.GetCurrentClassLogger();

            public static void OnChanged(object sender, FileSystemEventArgs e)
            {
                if (e.ChangeType != WatcherChangeTypes.Changed)
                {
                    return;
                }
                try
                {
                    var file_name = e.Name;
                    var mapper = new Mapper(e.FullPath);
                    var rows = mapper.Take<Equipment>("Лист1", 65535).ToList();
                    _equipments = rows.Select(r => r.Value).ToList();

                    log.Info(file_name + " успешно считан!");

                    foreach (var eqp in _equipments)
                    {
                        var client = new EqpIntegrationServiceClient();
                        client.Open();

                        foreach (var eqpuipment in _equipments)
                        {
                            var req = new EqpInsertUpdateRequest()
                            {
                                headerField = new EqpInsertUpdateRequestHeader() { fileIdField = 0 },
                                itemField = new EqpInsertUpdateRequestItem()
                                {
                                    extensionIdField = eqpuipment.MnfNum ?? string.Empty, //привязка по заводкому номеру
                                    mnfNumField = eqpuipment.MnfNum ?? string.Empty,
                                    itemNameField = eqpuipment.FullName ?? string.Empty, //обязательное поле
                                    typeNameField = eqpuipment.EqpTypeName ?? string.Empty, //обязательное поле
                                    kindField = eqpuipment.EqpKindName ?? string.Empty, //обязательное поле
                                    stateField = eqpuipment.EqpStateName ?? string.Empty, //обязательное поле
                                    invNumField = string.IsNullOrWhiteSpace(eqpuipment.InvNum) ? "-" : eqpuipment.InvNum, //обязательное поле
                                    legalBasisField = eqpuipment.LegalBasis ?? string.Empty, //обязательное поле
                                    checkTypeField = eqpuipment.EqpCheckTypeName ?? string.Empty, //обязательное поле
                                    locationField = eqpuipment.Location ?? string.Empty, //обязательное поле
                                    executerField = eqpuipment.UserId ?? string.Empty, //обязательное поле
                                    intervalTypeField = eqpuipment.IntervalTypeName ?? string.Empty, //обязательное поле
                                    intervalLenField = eqpuipment.IntervalLenStr.ToString() ?? string.Empty, //обязательное поле
                                    subdivisionField = eqpuipment.EqpSubDivisions ?? string.Empty, //обязательное поле
                                    inAccrScopeField = eqpuipment.InScopeAccreditation.ToString() ?? string.Empty, //обязательное поле
                                    mnfField = eqpuipment.Mnf ?? string.Empty,
                                    mnfDateField = eqpuipment.MnfDateStr.ToString("yyyy-MM-dd"),
                                    startUpField = eqpuipment.StartUpDateStr.ToString("yyyy-MM-dd"),
                                    checkDateField = eqpuipment.CheckDateStr.ToString("yyyy-MM-dd"),
                                    checkDocDateField = eqpuipment.CheckDateStr.ToString("yyyy-MM-dd"),
                                    checkNextDateField = eqpuipment.NextDateStr.ToString("yyyy-MM-dd"),
                                    checkPlaceField = eqpuipment.ECLPlace ?? string.Empty,
                                    checkDocField = eqpuipment.ECLDoc ?? string.Empty,
                                    //purposeField = eqpuipment.Purpose ?? string.Empty,
                                    itemNoteField = eqpuipment.Note ?? string.Empty,
                                    checkCommentField = eqpuipment.ECLComment ?? string.Empty,
                                }
                            };
                            try
                            {
                                var answer = client.InsUpd(req);
                                eqpuipment.Service_Response = answer.headerField.errMsgField;
                                log.Info("Оборудование: " + eqpuipment.FullName + " " + "Заводской номер: " + eqpuipment.MnfNum + " " + "Статус: " + eqpuipment.Service_Response);
                            }
                            catch (EndpointNotFoundException)
                            {
                                log.Error("Произошла ошибка: Нет связи с сервером\nУбедитесь в том, что:\n1. Серверный модуль интеграции I-LDS работает;\n2. Правильно указана строка подключения в файле: ImportEqpuipment.exe.config;\n3. Активность сервиса интеграции в файле: Indusoft.LDS.Server.DIM.Eqp.dll.config.");
                            }
                        }
                        client.Close();
                        File.Move(e.FullPath, e.FullPath + "_" + DateTime.Now.ToString("dd.MM.yyyy_HH.mm.ss"));
                        log.Info("Обработанный файл: " + e.Name + " " + "переименован: " + e.Name + "_" + DateTime.Now.ToString("dd.MM.yyyy_HH.mm.ss"));
                        Environment.Exit(1);
                    }
                }

                catch (Exception ex)
                {
                    log.Error("Ошибка: " + ex.Message);
                    Environment.Exit(1);
                }
            }
        }
    }
}