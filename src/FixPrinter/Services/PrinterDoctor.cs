using System.Diagnostics;
using System.Diagnostics.Eventing.Reader;
using System.IO;
using System.Management;
using System.Net.NetworkInformation;
using System.Net.Sockets;
using System.Printing;
using System.ServiceProcess;
using System.Text.Json;
using System.Text.Json.Nodes;
using System.Text.RegularExpressions;
using System.Windows;
using Microsoft.Win32;

namespace FixPrinter.Services;

public enum Severity { Ok, Warning, Error }

public sealed class CheckResult
{
    public Severity Severity { get; init; }
    public string Title { get; init; } = "";
    public string Details { get; init; } = "";
    /// <summary>Автоматическое исправление; null — чинится только вручную.</summary>
    public Action? Fix { get; init; }
    public bool IsProblem => Severity != Severity.Ok;
}

/// <summary>
/// Диагностика и починка подсистемы печати Windows:
/// служба печати, застрявшие задания, принтеры в «автономном режиме», сетевые порты.
/// </summary>
public sealed class PrinterDoctor
{
    private readonly Logger _log = Logger.Instance;

    private static readonly string[] VirtualPrinters =
    {
        "Microsoft Print to PDF", "Microsoft XPS Document Writer",
        "OneNote", "Fax", "AnyDesk Printer", "Adobe PDF", "Generic / Text Only"
    };

    /// <summary>Семейства аппаратов формата A3 — если драйвер такого принтера не отдаёт A3, драйвер неправильный.</summary>
    private static readonly string[] A3ClassModels =
    {
        "TASKalfa", "bizhub", "imageRUNNER", "WorkCentre", "AltaLink",
        "e-STUDIO", "ECOSYS M8", "ECOSYS M4", "Versant", "Aficio MP"
    };

    private static string SpoolDir =>
        Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows),
            "System32", "spool", "PRINTERS");

    // ---------- ДИАГНОСТИКА ----------

    /// <summary>Очереди, чей драйвер не ответил на опрос (виновники зависаний диалога печати).</summary>
    private readonly List<string> _hangingQueues = new();
    /// <summary>Принтеры, чьи драйверы поддерживают A3.</summary>
    private readonly List<string> _a3Capable = new();

    public List<CheckResult> Diagnose()
    {
        var results = new List<CheckResult>();
        _log.Info("--- Диагностика ---");
        _hangingQueues.Clear();
        _a3Capable.Clear();

        CheckService(results, "RpcSs", "Служба удалённого вызова процедур (RPC)", critical: true);
        CheckSpooler(results);
        CheckSpoolFolder(results);
        CheckWindowsManagesDefault(results);
        CheckPrinters(results);
        LogDriverVersions();
        CheckDuplicateQueues(results);
        CheckPaperSizesAndDefault(results);
        CheckBrowserLastPrinter(results);
        CheckAppCrashes(results);
        CheckPrintServiceLog(results);
        CheckStuckJobs(results);
        CheckNetworkPorts(results);

        foreach (var r in results)
            _log.Info($"[{r.Severity}] {r.Title}" + (r.Details.Length > 0 ? $" — {r.Details}" : ""));
        return results;
    }

    private void CheckService(List<CheckResult> results, string name, string title, bool critical)
    {
        try
        {
            using var sc = new ServiceController(name);
            if (sc.Status == ServiceControllerStatus.Running)
                results.Add(new CheckResult { Severity = Severity.Ok, Title = $"{title}: работает" });
            else
                results.Add(new CheckResult
                {
                    Severity = critical ? Severity.Error : Severity.Warning,
                    Title = $"{title}: не запущена (статус {sc.Status})",
                    Details = "Без этой службы печать невозможна.",
                    Fix = () => StartService(name)
                });
        }
        catch (Exception ex)
        {
            _log.Error($"Не удалось проверить службу {name}: {ex.Message}");
            results.Add(new CheckResult { Severity = Severity.Error, Title = $"{title}: не найдена", Details = ex.Message });
        }
    }

    private void CheckSpooler(List<CheckResult> results)
    {
        try
        {
            using var sc = new ServiceController("Spooler");
            var startMode = GetServiceStartMode("Spooler");

            if (startMode is "Disabled" or "Manual")
                results.Add(new CheckResult
                {
                    Severity = Severity.Error,
                    Title = $"Служба печати: автозапуск выключен ({startMode})",
                    Details = "Служба должна запускаться автоматически.",
                    Fix = () => SetServiceAutoStart("Spooler")
                });

            if (sc.Status != ServiceControllerStatus.Running)
                results.Add(new CheckResult
                {
                    Severity = Severity.Error,
                    Title = $"Служба печати (Spooler): не работает (статус {sc.Status})",
                    Fix = RestartSpoolerAndClearQueue
                });
            else
                results.Add(new CheckResult { Severity = Severity.Ok, Title = "Служба печати (Spooler): работает" });
        }
        catch (Exception ex)
        {
            _log.Error("Не удалось проверить Spooler: " + ex.Message);
            results.Add(new CheckResult { Severity = Severity.Error, Title = "Служба печати не найдена", Details = ex.Message });
        }
    }

    private void CheckSpoolFolder(List<CheckResult> results)
    {
        try
        {
            if (!Directory.Exists(SpoolDir))
            {
                results.Add(new CheckResult
                {
                    Severity = Severity.Error,
                    Title = "Папка очереди печати не существует",
                    Details = SpoolDir,
                    Fix = () => { Directory.CreateDirectory(SpoolDir); _log.Info("Папка очереди создана."); }
                });
                return;
            }

            var files = Directory.GetFiles(SpoolDir);
            var old = files.Count(f => DateTime.Now - File.GetLastWriteTime(f) > TimeSpan.FromMinutes(15));
            if (old > 0)
                results.Add(new CheckResult
                {
                    Severity = Severity.Warning,
                    Title = $"В очереди печати застряли файлы: {old} шт. старше 15 минут",
                    Details = "Очередь будет очищена с перезапуском службы печати.",
                    Fix = RestartSpoolerAndClearQueue
                });
            else
                results.Add(new CheckResult { Severity = Severity.Ok, Title = "Папка очереди печати: чисто" });
        }
        catch (Exception ex)
        {
            _log.Error("Проверка папки очереди: " + ex.Message);
        }
    }

    private void CheckPrinters(List<CheckResult> results)
    {
        try
        {
            using var searcher = new ManagementObjectSearcher("SELECT * FROM Win32_Printer");
            var printers = searcher.Get().Cast<ManagementObject>().ToList();

            var real = printers.Where(p => !IsVirtual((string)p["Name"])).ToList();
            _log.Info($"Найдено принтеров: {printers.Count} (из них реальных: {real.Count})");

            if (real.Count == 0)
                results.Add(new CheckResult
                {
                    Severity = Severity.Warning,
                    Title = "Не найдено ни одного реального принтера",
                    Details = "Установлены только виртуальные (PDF/XPS). Если принтер подключён — " +
                              "проверьте кабель/сеть и установите драйвер с сайта производителя."
                });

            if (!printers.Any(p => (bool)(p["Default"] ?? false)))
                results.Add(new CheckResult
                {
                    Severity = Severity.Warning,
                    Title = "Не назначен принтер по умолчанию",
                    Details = "Параметры → Bluetooth и устройства → Принтеры и сканеры."
                });

            foreach (var p in printers)
            {
                var name = (string)p["Name"];
                var offline = (bool)(p["WorkOffline"] ?? false);
                var status = Convert.ToInt32(p["PrinterStatus"] ?? 0);
                var errState = Convert.ToInt32(p["DetectedErrorState"] ?? 0);
                var driver = (string)(p["DriverName"] ?? "?");
                var port = (string)(p["PortName"] ?? "?");
                var isDefault = (bool)(p["Default"] ?? false);

                _log.Info($"Принтер «{name}»: драйвер «{driver}», порт «{port}», " +
                          $"по умолчанию: {(isDefault ? "ДА" : "нет")}, offline: {offline}, " +
                          $"статус: {status}, ошибка: {errState}");

                if (!IsVirtual(name) && port.StartsWith("WSD", StringComparison.OrdinalIgnoreCase))
                    results.Add(new CheckResult
                    {
                        Severity = Severity.Warning,
                        Title = $"«{name}» подключён через WSD-порт",
                        Details = "Когда такой принтер выключен или «уснул», диалоги печати в Chrome, " +
                                  "Acrobat и других программах подолгу зависают. Включите принтер. " +
                                  "Надёжное решение — переустановить принтер через порт TCP/IP (по IP-адресу)."
                    });

                if (offline && !IsVirtual(name))
                {
                    results.Add(new CheckResult
                    {
                        Severity = Severity.Error,
                        Title = $"«{name}»: включён автономный режим (offline)",
                        Details = "Windows не отправляет задания на этот принтер.",
                        Fix = () => SetPrinterOnline(name)
                    });
                }
                else if (errState is 8 or 7 or 4 or 6) // замятие, открыта крышка, нет бумаги, нет тонера
                {
                    var reason = errState switch
                    {
                        4 => "нет бумаги", 6 => "закончился тонер/чернила",
                        7 => "открыта крышка", 8 => "замятие бумаги", _ => "ошибка"
                    };
                    results.Add(new CheckResult
                    {
                        Severity = Severity.Error,
                        Title = $"«{name}»: {reason}",
                        Details = "Это аппаратная проблема — устраните её на самом принтере."
                    });
                }
                else if (status == 7 && !IsVirtual(name))
                {
                    results.Add(new CheckResult
                    {
                        Severity = Severity.Warning,
                        Title = $"«{name}»: принтер сообщает статус Offline",
                        Details = "Проверьте питание и подключение. Попробую перевести в онлайн.",
                        Fix = () => SetPrinterOnline(name)
                    });
                }
                else
                {
                    results.Add(new CheckResult { Severity = Severity.Ok, Title = $"«{name}»: в порядке" });
                }
            }
        }
        catch (Exception ex)
        {
            _log.Error("Проверка принтеров: " + ex);
            results.Add(new CheckResult { Severity = Severity.Error, Title = "Не удалось опросить принтеры (WMI)", Details = ex.Message });
        }
    }

    private sealed record PrinterRow(string Name, string Driver, string Port, bool Default);

    private List<PrinterRow> ListPrinterRows()
    {
        using var searcher = new ManagementObjectSearcher(
            "SELECT Name, DriverName, PortName, Default FROM Win32_Printer");
        return searcher.Get().Cast<ManagementObject>()
            .Select(p => new PrinterRow(
                (string)p["Name"],
                (string)(p["DriverName"] ?? "?"),
                (string)(p["PortName"] ?? "?"),
                (bool)(p["Default"] ?? false)))
            .ToList();
    }

    /// <summary>Windows 10/11 по умолчанию сама переключает принтер «по умолчанию» на последний использованный —
    /// из-за этого печать внезапно уходит не туда. Отключаем это поведение.</summary>
    private void CheckWindowsManagesDefault(List<CheckResult> results)
    {
        try
        {
            using var key = Registry.CurrentUser.OpenSubKey(
                @"Software\Microsoft\Windows NT\CurrentVersion\Windows");
            var mode = key?.GetValue("LegacyDefaultPrinterMode") as int?;
            if (mode == 1)
            {
                results.Add(new CheckResult { Severity = Severity.Ok, Title = "Windows не переключает принтер по умолчанию сама" });
                return;
            }
            results.Add(new CheckResult
            {
                Severity = Severity.Warning,
                Title = "Windows сама переключает принтер по умолчанию",
                Details = "Включено «Разрешить Windows управлять принтером по умолчанию»: печать может " +
                          "внезапно уходить на другой (в т.ч. виртуальный) принтер. Отключу это.",
                Fix = () =>
                {
                    using var wk = Registry.CurrentUser.CreateSubKey(
                        @"Software\Microsoft\Windows NT\CurrentVersion\Windows");
                    wk.SetValue("LegacyDefaultPrinterMode", 1, RegistryValueKind.DWord);
                    _log.Info("LegacyDefaultPrinterMode = 1: автопереключение принтера по умолчанию отключено.");
                }
            });
        }
        catch (Exception ex)
        {
            _log.Warn("Проверка LegacyDefaultPrinterMode: " + ex.Message);
        }
    }

    /// <summary>Очереди «(копия 1)» — следы переустановки драйвера. Если оригинал жив — копию можно удалить.</summary>
    private void CheckDuplicateQueues(List<CheckResult> results)
    {
        try
        {
            var printers = ListPrinterRows();
            var rx = new Regex(@"^(?<base>.+?)\s*\((копия|copy)\s*\d+\)\s*$", RegexOptions.IgnoreCase);
            foreach (var p in printers)
            {
                var m = rx.Match(p.Name);
                if (!m.Success) continue;
                var baseName = m.Groups["base"].Value.Trim();
                var original = printers.FirstOrDefault(o =>
                    o.Name.Equals(baseName, StringComparison.OrdinalIgnoreCase));

                if (original != null)
                {
                    var copyName = p.Name;
                    results.Add(new CheckResult
                    {
                        Severity = Severity.Warning,
                        Title = $"Дубликат очереди печати: «{copyName}»",
                        Details = $"Есть и оригинал «{baseName}», и копия — программы путаются, " +
                                  "какая очередь рабочая. Удалю копию.",
                        Fix = () => DeletePrinterQueue(copyName)
                    });
                }
                else
                {
                    _log.Info($"Очередь «{p.Name}» названа «копия», но оригинала нет — работе не мешает.");
                }
            }
        }
        catch (Exception ex)
        {
            _log.Warn("Проверка дубликатов очередей: " + ex.Message);
        }
    }

    /// <summary>Опрашивает у драйверов поддерживаемые форматы бумаги (это же выявляет зависшие драйверы —
    /// главную причину тормозов диалога печати) и проверяет принтер по умолчанию.</summary>
    private void CheckPaperSizesAndDefault(List<CheckResult> results)
    {
        List<PrinterRow> printers;
        try { printers = ListPrinterRows(); }
        catch (Exception ex) { _log.Error("Список принтеров: " + ex.Message); return; }

        var real = printers.Where(p => !IsVirtual(p.Name)).ToList();
        var a3Capable = new List<string>();

        foreach (var p in real)
        {
            var sizes = new List<string>();
            var sw = Stopwatch.StartNew();
            try
            {
                var queueName = p.Name;
                var task = Task.Run(() =>
                {
                    using var server = new LocalPrintServer();
                    using var q = new PrintQueue(server, queueName);
                    var caps = q.GetPrintCapabilities();
                    return caps.PageMediaSizeCapability
                        .Where(s => s.PageMediaSizeName.HasValue)
                        .Select(s => s.PageMediaSizeName!.Value.ToString())
                        .ToList();
                });

                if (!task.Wait(TimeSpan.FromSeconds(12)))
                {
                    _hangingQueues.Add(p.Name);
                    results.Add(BuildHangingQueueIssue(p));
                    continue;
                }
                sizes = task.Result;
            }
            catch (Exception ex)
            {
                _log.Warn($"Не удалось получить форматы бумаги «{p.Name}»: {ex.InnerException?.Message ?? ex.Message}");
                continue;
            }

            _log.Info($"«{p.Name}»: драйвер ответил за {sw.ElapsedMilliseconds} мс, " +
                      $"форматов бумаги: {sizes.Count} — {string.Join(", ", sizes)}");
            if (sw.ElapsedMilliseconds > 5000)
                results.Add(new CheckResult
                {
                    Severity = Severity.Warning,
                    Title = $"Драйвер «{p.Name}» отвечает очень медленно ({sw.ElapsedMilliseconds / 1000.0:F1} с)",
                    Details = "Это замедляет открытие диалога печати в любых программах."
                });

            var hasA3 = sizes.Contains("ISOA3");
            if (hasA3)
            {
                a3Capable.Add(p.Name);
            }
            else if (A3ClassModels.Any(m2 => p.Name.Contains(m2, StringComparison.OrdinalIgnoreCase) ||
                                             p.Driver.Contains(m2, StringComparison.OrdinalIgnoreCase)))
            {
                results.Add(new CheckResult
                {
                    Severity = Severity.Error,
                    Title = $"«{p.Name}» — аппарат формата A3, но его драйвер не предлагает A3",
                    Details = "Поэтому формата A3 нет в Chrome/Acrobat. Установлен урезанный или неверный " +
                              "драйвер — скачайте полный драйвер именно вашей модели с сайта производителя " +
                              "и переустановите принтер."
                });
            }
        }

        if (a3Capable.Count > 0)
            results.Add(new CheckResult
            {
                Severity = Severity.Ok,
                Title = "Печать A3 доступна на: " + string.Join(", ", a3Capable.Select(n => $"«{n}»"))
            });
        else if (real.Count > 0)
            results.Add(new CheckResult
            {
                Severity = Severity.Warning,
                Title = "Ни один установленный драйвер не предлагает формат A3",
                Details = "Если принтер физически умеет A3 — переустановите его полный драйвер. " +
                          "Если все принтеры формата A4 — A3 на них не напечатать."
            });

        _a3Capable.AddRange(a3Capable);

        // Принтер по умолчанию
        var def = printers.FirstOrDefault(p => p.Default);
        _log.Info("Принтер по умолчанию: " + (def == null ? "НЕ НАЗНАЧЕН" : $"«{def.Name}»"));
        if (def != null && IsVirtual(def.Name) && real.Count > 0)
        {
            var candidate = a3Capable.FirstOrDefault() ?? real[0].Name;
            results.Add(new CheckResult
            {
                Severity = Severity.Warning,
                Title = $"Принтер по умолчанию — виртуальный: «{def.Name}»",
                Details = $"Программы при печати целятся в него, а не в настоящий принтер. " +
                          $"Сделаю принтером по умолчанию «{candidate}».",
                Fix = () => SetDefaultPrinter(candidate)
            });
        }
        else if (def != null && a3Capable.Count > 0 && !a3Capable.Contains(def.Name))
        {
            results.Add(new CheckResult
            {
                Severity = Severity.Warning,
                Title = $"Принтер по умолчанию «{def.Name}» не умеет A3",
                Details = "Для печати A3 выбирайте в диалоге печати: " +
                          string.Join(", ", a3Capable.Select(n => $"«{n}»"))
            });
        }
    }

    /// <summary>Разбирается, почему очередь зависла: ищет IP WSD-устройства в реестре,
    /// проверяет доступность по сети и подбирает подходящее лечение.</summary>
    private CheckResult BuildHangingQueueIssue(PrinterRow p)
    {
        var isWsd = p.Port.StartsWith("WSD", StringComparison.OrdinalIgnoreCase);
        string? host = null;
        if (isWsd)
        {
            host = ResolveWsdDeviceHost(p.Name);
            _log.Info($"WSD-устройство «{p.Name}»: адрес из реестра — {(host ?? "не найден")}");
        }

        if (host != null)
        {
            var reachable = IsDeviceReachable(host);
            _log.Info($"«{p.Name}» ({host}): {(reachable ? "ДОСТУПЕН в сети" : "НЕДОСТУПЕН в сети")}");

            if (reachable)
                return new CheckResult
                {
                    Severity = Severity.Error,
                    Title = $"Драйвер «{p.Name}» завис, хотя сам принтер в сети ({host})",
                    Details = "Перезапущу службу печати — обычно это снимает зависание. " +
                              "Если повторится — переустановите принтер через порт TCP/IP с драйвером Epson.",
                    Fix = RestartSpoolerAndClearQueue
                };

            return new CheckResult
            {
                Severity = Severity.Error,
                Title = $"«{p.Name}» выключен или недоступен в сети ({host}) — его очередь вешает диалоги печати",
                Details = "Из-за этого Chrome и Acrobat тормозят при печати. Варианты: включить принтер — " +
                          "или удалить его очередь (предложу подтверждение); когда принтер снова понадобится, " +
                          "Windows добавит его заново (Параметры → Принтеры → Добавить устройство).",
                Fix = () => ConfirmAndRemovePrinter(p.Name,
                    "сейчас выключен или недоступен в сети, а его очередь печати не отвечает и подвешивает " +
                    "диалоги печати в Chrome, Acrobat и других программах.")
            };
        }

        return new CheckResult
        {
            Severity = Severity.Error,
            Title = $"Драйвер принтера «{p.Name}» не отвечает (ждал 12 секунд)",
            Details = "Именно такие зависания тормозят диалоги печати в Chrome и Acrobat. " +
                      "Предложу удалить эту очередь печати (принтер можно добавить заново в любой момент).",
            Fix = () => ConfirmAndRemovePrinter(p.Name,
                "не отвечает: его очередь подвешивает диалоги печати в Chrome, Acrobat и других программах.")
        };
    }

    /// <summary>Ищет адрес WSD-устройства в реестре (SWD\DAFWSDProvider → LocationInformation).</summary>
    private string? ResolveWsdDeviceHost(string printerName)
    {
        try
        {
            using var root = Registry.LocalMachine.OpenSubKey(
                @"SYSTEM\CurrentControlSet\Enum\SWD\DAFWSDProvider");
            if (root == null) return null;

            foreach (var sub in root.GetSubKeyNames())
            {
                using var dev = root.OpenSubKey(sub);
                if (dev == null) continue;
                var friendly = dev.GetValue("FriendlyName") as string ?? "";
                var location = dev.GetValue("LocationInformation") as string ?? "";
                if (friendly.Length == 0 || location.Length == 0) continue;

                // Имя принтера обычно совпадает с FriendlyName устройства (или начинается с него)
                if (printerName.Contains(friendly, StringComparison.OrdinalIgnoreCase) ||
                    friendly.Contains(printerName, StringComparison.OrdinalIgnoreCase) ||
                    printerName.Split(' ')[0].Equals(friendly.Split(' ')[0], StringComparison.OrdinalIgnoreCase))
                {
                    if (Uri.TryCreate(location, UriKind.Absolute, out var uri))
                        return uri.Host;
                }
            }
        }
        catch (Exception ex)
        {
            _log.Warn("Поиск адреса WSD-устройства: " + ex.Message);
        }
        return null;
    }

    private bool IsDeviceReachable(string host)
    {
        try
        {
            using var ping = new Ping();
            if (ping.Send(host, 1200).Status == IPStatus.Success) return true;
        }
        catch { }
        // Пинг может быть закрыт — пробуем типичные порты печати
        foreach (var port in new[] { 9100, 631, 80, 515 })
            if (TcpReachable(host, port, 1200)) return true;
        return false;
    }

    private void ConfirmAndRemovePrinter(string printerName, string reason)
    {
        var answer = MessageBox.Show(
            $"Принтер «{printerName}» {reason}\n\n" +
            "Удалить эту очередь печати? Сам принтер не пострадает — когда он понадобится, " +
            "его можно добавить заново (Параметры → Bluetooth и устройства → Принтеры).",
            "FixPrinter — нужно ваше решение",
            MessageBoxButton.YesNo, MessageBoxImage.Question);

        if (answer != MessageBoxResult.Yes)
        {
            _log.Info($"Пользователь отказался удалять очередь «{printerName}».");
            return;
        }
        DeletePrinterQueue(printerName);
    }

    /// <summary>Chrome и Edge открывают печать на последнем использованном принтере.
    /// Если это зависшая или не умеющая A3 очередь — предупреждаем.</summary>
    private void CheckBrowserLastPrinter(List<CheckResult> results)
    {
        List<string> existing;
        try { existing = ListPrinterRows().Select(r => r.Name).ToList(); }
        catch { existing = new List<string>(); }

        var local = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
        var browsers = new (string Browser, string Proc, string UserData)[]
        {
            ("Chrome", "chrome", Path.Combine(local, "Google", "Chrome", "User Data")),
            ("Edge",   "msedge", Path.Combine(local, "Microsoft", "Edge", "User Data"))
        };

        foreach (var (browser, proc, userData) in browsers)
        {
            try
            {
                if (!Directory.Exists(userData)) continue;
                foreach (var profileDir in Directory.GetDirectories(userData)
                             .Where(d => Path.GetFileName(d) == "Default" ||
                                         Path.GetFileName(d).StartsWith("Profile ")))
                {
                    var prefs = Path.Combine(profileDir, "Preferences");
                    if (!File.Exists(prefs)) continue;
                    var lastPrinter = ReadLastPrinterFromPrefs(prefs);
                    if (lastPrinter == null) continue;

                    var profile = Path.GetFileName(profileDir);
                    _log.Info($"{browser} ({profile}): последний использованный принтер — «{lastPrinter}»");

                    if (lastPrinter != "Save as PDF" && !existing.Contains(lastPrinter))
                        results.Add(new CheckResult
                        {
                            Severity = Severity.Error,
                            Title = $"{browser} запомнил принтер «{lastPrinter}», которого уже нет",
                            Details = $"Печать в {browser} из-за этого может зависать или ронять браузер. " +
                                      $"Сброшу запомненные настройки печати {browser} (нужно, чтобы {browser} был закрыт).",
                            Fix = () => CleanBrowserPrintState(browser, proc, prefs)
                        });
                    else if (_hangingQueues.Contains(lastPrinter))
                        results.Add(new CheckResult
                        {
                            Severity = Severity.Error,
                            Title = $"{browser} запомнил зависший принтер «{lastPrinter}» как последний",
                            Details = $"Поэтому печать в {browser} тормозит с порога: диалог сразу опрашивает " +
                                      "эту очередь. Откройте печать (Ctrl+P), дождитесь списка и выберите " +
                                      "другой принтер — браузер запомнит новый выбор."
                        });
                    else if (_a3Capable.Count > 0 && !_a3Capable.Contains(lastPrinter) &&
                             !IsVirtual(lastPrinter) && lastPrinter != "Save as PDF")
                        results.Add(new CheckResult
                        {
                            Severity = Severity.Warning,
                            Title = $"{browser} печатает на «{lastPrinter}», а он не умеет A3",
                            Details = "В диалоге печати (Ctrl+P) выберите: " +
                                      string.Join(", ", _a3Capable.Select(n => $"«{n}»")) +
                                      " — там A3 есть, и браузер запомнит выбор."
                        });
                }
            }
            catch (Exception ex)
            {
                _log.Warn($"Проверка последнего принтера {browser}: {ex.Message}");
            }
        }
    }

    /// <summary>Сбрасывает запомненные настройки печати Chromium-браузера (с резервной копией).</summary>
    private void CleanBrowserPrintState(string browserTitle, string processName, string prefsPath)
    {
        if (Process.GetProcessesByName(processName).Length > 0)
        {
            _log.Warn($"{browserTitle} сейчас запущен — настройки печати не сбросить, пока он открыт.");
            MessageBox.Show(
                $"Полностью закройте {browserTitle} и нажмите «Найти проблемы и починить» ещё раз — " +
                "тогда я сброшу его запомненные настройки печати.",
                "FixPrinter", MessageBoxButton.OK, MessageBoxImage.Information);
            return;
        }

        File.Copy(prefsPath, prefsPath + ".fixprinter.bak", overwrite: true);
        var root = JsonNode.Parse(File.ReadAllText(prefsPath))!.AsObject();
        if (root["printing"] is JsonObject printing)
        {
            printing.Remove("print_preview_sticky_settings");
            if (printing.Count == 0) root.Remove("printing");
        }
        File.WriteAllText(prefsPath, root.ToJsonString());
        _log.Info($"{browserTitle}: запомненные настройки печати сброшены " +
                  $"(резервная копия: {Path.GetFileName(prefsPath)}.fixprinter.bak). " +
                  "При следующей печати браузер предложит принтер по умолчанию.");
    }

    private static readonly string[] PrintCrashApps =
        { "chrome.exe", "msedge.exe", "acrord32.exe", "acrobat.exe", "splwow64.exe", "spoolsv.exe" };

    /// <summary>Ищет в журнале Windows падения браузеров/Acrobat/служб печати (Event 1000)
    /// и определяет сбойный модуль. Если модуль принадлежит драйверу принтера — чинит.</summary>
    private void CheckAppCrashes(List<CheckResult> results)
    {
        try
        {
            const long twoWeeksMs = 14L * 24 * 3600 * 1000;
            var query = new EventLogQuery("Application", PathType.LogName,
                $"*[System[Provider[@Name='Application Error'] and (EventID=1000) " +
                $"and TimeCreated[timediff(@SystemTime) <= {twoWeeksMs}]]]");
            using var reader = new EventLogReader(query);

            var crashes = new List<(DateTime Time, string App, string Module, string ModulePath)>();
            for (EventRecord? rec; (rec = reader.ReadEvent()) != null; rec.Dispose())
            {
                var p = rec.Properties;
                string S(int i) => p.Count > i ? p[i].Value?.ToString() ?? "" : "";
                var app = S(0);
                if (!PrintCrashApps.Contains(app, StringComparer.OrdinalIgnoreCase)) continue;
                crashes.Add((rec.TimeCreated ?? DateTime.MinValue, app, S(3), S(11)));
            }

            if (crashes.Count == 0)
            {
                _log.Info("Падений браузеров/Acrobat/службы печати за 14 дней в журнале Windows нет.");
                results.Add(new CheckResult { Severity = Severity.Ok, Title = "Аварийных завершений при печати в журнале Windows нет" });
                return;
            }

            foreach (var c in crashes.OrderBy(c => c.Time))
                _log.Info($"Журнал Windows: {c.App} упал {c.Time:dd.MM HH:mm}, модуль: {c.Module} ({c.ModulePath})");

            // Падения самой подсистемы печати
            foreach (var group in crashes
                         .Where(c => c.App.Equals("spoolsv.exe", StringComparison.OrdinalIgnoreCase) ||
                                     c.App.Equals("splwow64.exe", StringComparison.OrdinalIgnoreCase))
                         .GroupBy(c => c.App, StringComparer.OrdinalIgnoreCase))
            {
                var last = group.OrderByDescending(c => c.Time).First();
                var guilty = QueuesUsingDriverFile(last.Module);
                results.Add(new CheckResult
                {
                    Severity = Severity.Error,
                    Title = $"Служба печати аварийно завершалась ({group.Key}, {group.Count()} раз, модуль {last.Module})",
                    Details = "Перезапущу службу печати и очищу очередь." +
                              (guilty.Count > 0
                                  ? $" Сбойный модуль принадлежит драйверу: {string.Join(", ", guilty.Select(q => $"«{q}»"))}. " +
                                    DriverDownloadHint(guilty)
                                  : ""),
                    Fix = RestartSpoolerAndClearQueue
                });
            }

            // Падения приложений (Chrome/Edge/Acrobat) — группируем по сбойному модулю,
            // чтобы не плодить одинаковые проблемы и не чинить одно и то же дважды
            var appCrashes = crashes
                .Where(c => !c.App.Equals("spoolsv.exe", StringComparison.OrdinalIgnoreCase) &&
                            !c.App.Equals("splwow64.exe", StringComparison.OrdinalIgnoreCase))
                .ToList();

            foreach (var moduleGroup in appCrashes.GroupBy(c => c.Module, StringComparer.OrdinalIgnoreCase))
            {
                var last = moduleGroup.OrderByDescending(c => c.Time).First();
                var apps = string.Join(", ", moduleGroup
                    .GroupBy(c => c.App, StringComparer.OrdinalIgnoreCase)
                    .Select(g => $"{AppTitle(g.Key)} ({g.Count()} раз)"));

                var guiltyQueues = QueuesUsingDriverFile(last.Module);
                var inPrintDriver = guiltyQueues.Count > 0 ||
                                    last.ModulePath.Contains(@"\spool\", StringComparison.OrdinalIgnoreCase);

                if (!inPrintDriver)
                {
                    results.Add(new CheckResult
                    {
                        Severity = Severity.Warning,
                        Title = $"Аварийные завершения: {apps} (модуль {last.Module})",
                        Details = "Сбойный модуль не похож на драйвер принтера. Если браузер закрывается именно " +
                                  "при нажатии «Печать» — пришлите логи после такого случая, разберу подробнее."
                    });
                    continue;
                }

                var target = guiltyQueues.Count > 0 ? guiltyQueues : AllRealQueues();
                var treatment = GetTreatment(target);

                if (treatment == null)
                {
                    results.Add(new CheckResult
                    {
                        Severity = Severity.Error,
                        Title = $"Падения при печати из-за драйвера принтера ({last.Module}): {apps}",
                        Details = "Виновник: " + string.Join(", ", target.Select(q => $"«{q}»")) +
                                  ". Отключу у него «расширенные функции печати» (печать в RAW) и перезапущу " +
                                  "службу печати. Обходной путь: в Chrome печатайте через системный " +
                                  "диалог (Ctrl+Shift+P).",
                        Fix = () =>
                        {
                            foreach (var q in target) SetRawOnly(q, true);
                            RestartSpoolerAndClearQueue();
                        }
                    });
                    continue;
                }

                var after = moduleGroup.Count(c => c.Time > treatment.AppliedAt.AddMinutes(2));
                if (after == 0)
                {
                    results.Add(new CheckResult
                    {
                        Severity = Severity.Ok,
                        Title = $"Падения из-за {last.Module} были до лечения ({treatment.AppliedAt:dd.MM HH:mm}) — новых нет",
                        Details = "Проверьте печать; если браузер снова закроется — запустите проверку ещё раз."
                    });
                }
                else if (treatment.Stage == 1)
                {
                    // RawOnly не помог: крэш происходит при загрузке конфигурационной DLL драйвера
                    // в процессе браузера. Этап 2: сброс повреждённого DEVMODE + пересоздание очереди.
                    results.Add(new CheckResult
                    {
                        Severity = Severity.Error,
                        Title = $"Браузер падает из-за драйвера ({last.Module}) даже после первого лечения ({after} раз)",
                        Details = "Применю лечение, этап 2: сброшу сохранённые настройки печати (повреждённый " +
                                  "DEVMODE в реестре — типовая причина таких падений; сделаю резервную копию) и " +
                                  "пересоздам очередь печати начисто. " +
                                  "До проверки печатайте через системный диалог (Ctrl+Shift+P).",
                        Fix = () =>
                        {
                            foreach (var q in target)
                            {
                                ClearPerUserDevMode(q);
                                var newName = RecreateQueueClean(q);
                                MarkTreatment(newName ?? q, 2);
                                if (newName != null) MarkTreatment(q, 2);
                            }
                            RestartSpoolerAndClearQueue();
                        }
                    });
                }
                else
                {
                    var hint = DriverDownloadHint(target);
                    results.Add(new CheckResult
                    {
                        Severity = Severity.Error,
                        Title = $"Браузер падает из-за драйвера ({last.Module}) после двух этапов лечения",
                        Details = "Автоматические меры исчерпаны: драйвер повреждён или несовместим — нужна " +
                                  "чистая переустановка. Открою страницу загрузки. Порядок: 1) скачайте драйвер; " +
                                  "2) в FixPrinter нажмите «Удалить принтеры…» и удалите этот принтер вместе с " +
                                  "драйвером; 3) установите скачанный драйвер и добавьте принтер заново. " + hint,
                        Fix = () =>
                        {
                            var url = Regex.Match(hint, @"https?://\S+").Value;
                            if (url.Length > 0)
                            {
                                _log.Info("Открываю страницу загрузки драйвера: " + url);
                                Process.Start(new ProcessStartInfo(url) { UseShellExecute = true });
                            }
                        }
                    });
                }
            }
        }
        catch (Exception ex)
        {
            _log.Warn("Чтение журнала падений: " + ex.Message);
        }
    }

    /// <summary>Ошибки журнала Microsoft-Windows-PrintService/Admin за 14 дней.</summary>
    private void CheckPrintServiceLog(List<CheckResult> results)
    {
        try
        {
            const long twoWeeksMs = 14L * 24 * 3600 * 1000;
            var query = new EventLogQuery("Microsoft-Windows-PrintService/Admin", PathType.LogName,
                $"*[System[(Level=2) and TimeCreated[timediff(@SystemTime) <= {twoWeeksMs}]]]");
            using var reader = new EventLogReader(query);

            var messages = new List<string>();
            var count = 0;
            for (EventRecord? rec; (rec = reader.ReadEvent()) != null; rec.Dispose())
            {
                count++;
                if (messages.Count < 5)
                {
                    string msg;
                    try { msg = rec.FormatDescription() ?? $"EventID {rec.Id}"; }
                    catch { msg = $"EventID {rec.Id}"; }
                    msg = Regex.Replace(msg, @"\s+", " ").Trim();
                    messages.Add($"{rec.TimeCreated:dd.MM HH:mm} — {(msg.Length > 220 ? msg[..220] + "…" : msg)}");
                }
            }

            foreach (var m in messages) _log.Info("Журнал печати: " + m);
            if (count > 0)
                results.Add(new CheckResult
                {
                    Severity = Severity.Warning,
                    Title = $"В журнале службы печати есть ошибки: {count} за 14 дней",
                    Details = messages.FirstOrDefault() ?? ""
                });
            else
                results.Add(new CheckResult { Severity = Severity.Ok, Title = "Журнал службы печати: ошибок за 14 дней нет" });
        }
        catch (Exception ex)
        {
            _log.Info("Журнал PrintService/Admin недоступен (обычно он выключен по умолчанию): " + ex.Message);
        }
    }

    private static string AppTitle(string exe) =>
        exe.StartsWith("chrome", StringComparison.OrdinalIgnoreCase) ? "Chrome"
        : exe.StartsWith("msedge", StringComparison.OrdinalIgnoreCase) ? "Edge"
        : exe.StartsWith("acro", StringComparison.OrdinalIgnoreCase) ? "Adobe Acrobat"
        : exe;

    /// <summary>Подсказка, где взять свежий драйвер для конкретных очередей.</summary>
    private string DriverDownloadHint(IEnumerable<string> queueNames)
    {
        var all = string.Join(" ", queueNames);
        string driver = "";
        try
        {
            var rows = ListPrinterRows();
            driver = string.Join(" ", rows.Where(r => queueNames.Contains(r.Name)).Select(r => r.Driver));
        }
        catch { }
        var hay = all + " " + driver;

        if (hay.Contains("TASKalfa 180", StringComparison.OrdinalIgnoreCase) ||
            hay.Contains("TASKalfa 220", StringComparison.OrdinalIgnoreCase))
            return "Последний официальный драйвер — GX Driver 5.0.2906, страница загрузки Kyocera: " +
                   "https://www.kyoceradocumentsolutions.eu/en/support/downloads.name-L2V1L2VuL21mcC9UQVNLQUxGQTE4MA==.html " +
                   "— скачайте GX Driver, удалите текущий драйвер (кнопка «Удалить принтеры…» в FixPrinter) " +
                   "и установите заново.";
        if (hay.Contains("Kyocera", StringComparison.OrdinalIgnoreCase))
            return "Свежий драйвер: https://www.kyoceradocumentsolutions.eu/en/support/downloads.html (выберите свою модель).";
        if (hay.Contains("EPSON", StringComparison.OrdinalIgnoreCase))
            return "Свежий драйвер: https://www.epson.ru/support (выберите свою модель).";
        return "Установите свежий драйвер с официального сайта производителя принтера.";
    }

    // ---------- Память о применённых лечениях ----------

    private static string StateFilePath => Path.Combine(
        Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
        "FixPrinter", "state.json");

    /// <summary>Этапы лечения падений из-за драйвера: 1 — RawOnly, 2 — сброс DEVMODE + пересоздание очереди.</summary>
    public sealed record Treatment(int Stage, DateTime AppliedAt);

    private Dictionary<string, Treatment> LoadTreatments()
    {
        try
        {
            if (File.Exists(StateFilePath))
            {
                var text = File.ReadAllText(StateFilePath);
                try
                {
                    var t = JsonSerializer.Deserialize<Dictionary<string, Treatment>>(text);
                    if (t != null) return t;
                }
                catch { /* старый формат */ }
                var old = JsonSerializer.Deserialize<Dictionary<string, DateTime>>(text);
                if (old != null)
                    return old.ToDictionary(kv => kv.Key, kv => new Treatment(1, kv.Value));
            }
        }
        catch (Exception ex) { _log.Warn("Чтение state.json: " + ex.Message); }
        return new Dictionary<string, Treatment>();
    }

    private void MarkTreatment(string printerName, int stage)
    {
        try
        {
            var state = LoadTreatments();
            state[printerName] = new Treatment(stage, DateTime.Now);
            Directory.CreateDirectory(Path.GetDirectoryName(StateFilePath)!);
            File.WriteAllText(StateFilePath, JsonSerializer.Serialize(state));
            _log.Info($"Отмечено лечение: «{printerName}», этап {stage}.");
        }
        catch (Exception ex) { _log.Warn("Запись state.json: " + ex.Message); }
    }

    /// <summary>Последнее применённое лечение для этих очередей (null — не применяли).</summary>
    private Treatment? GetTreatment(IEnumerable<string> queueNames)
    {
        var state = LoadTreatments();
        Treatment? best = null;
        foreach (var name in queueNames)
            if (state.TryGetValue(name, out var t) && (best == null || t.AppliedAt > best.AppliedAt))
                best = t;
        return best;
    }

    /// <summary>Сбрасывает пользовательские настройки печати (DEVMODE) принтера в реестре —
    /// повреждённый DEVMODE — типовая причина падений конфигурационной DLL драйвера
    /// внутри Chrome/Edge/Acrobat. Значения сохраняются в резервную копию.</summary>
    private void ClearPerUserDevMode(string printerName)
    {
        var backupDir = Path.Combine(Path.GetDirectoryName(StateFilePath)!, "devmode-backup");
        Directory.CreateDirectory(backupDir);

        foreach (var subKey in new[] { @"Printers\DevModePerUser", @"Printers\DevModes2" })
        {
            try
            {
                using var key = Registry.CurrentUser.OpenSubKey(subKey, writable: true);
                if (key == null) continue;
                foreach (var valueName in key.GetValueNames())
                {
                    if (!valueName.Equals(printerName, StringComparison.OrdinalIgnoreCase) &&
                        !valueName.StartsWith(printerName + ",", StringComparison.OrdinalIgnoreCase))
                        continue;

                    if (key.GetValue(valueName) is byte[] bytes)
                    {
                        var safe = Regex.Replace(valueName, @"[^\w\-]+", "_");
                        File.WriteAllBytes(Path.Combine(backupDir,
                            $"{safe}_{DateTime.Now:yyyyMMdd_HHmmss}.bin"), bytes);
                    }
                    key.DeleteValue(valueName);
                    _log.Info($"Сброшен сохранённый DEVMODE: {subKey}\\{valueName} (резервная копия в {backupDir}).");
                }
            }
            catch (Exception ex)
            {
                _log.Warn($"Сброс DEVMODE в {subKey}: {ex.Message}");
            }
        }
    }

    /// <summary>Пересоздаёт очередь с чистым именем (без «копия N») на том же драйвере и порту.
    /// Старая очередь удаляется только после успешного создания новой. Возвращает новое имя или null.</summary>
    private string? RecreateQueueClean(string oldName)
    {
        try
        {
            var rows = ListPrinterRows();
            var row = rows.FirstOrDefault(r => r.Name.Equals(oldName, StringComparison.OrdinalIgnoreCase));
            if (row == null) return null;

            var clean = Regex.Replace(oldName, @"\s*\((копия|copy)\s*\d+\)\s*$", "",
                RegexOptions.IgnoreCase).Trim();
            if (clean.Length == 0 || clean.Equals(oldName, StringComparison.OrdinalIgnoreCase) ||
                rows.Any(r => r.Name.Equals(clean, StringComparison.OrdinalIgnoreCase)))
                return null; // пересоздавать не во что — ограничиваемся сбросом DEVMODE

            _log.Info($"Пересоздаю очередь: «{oldName}» → «{clean}» (драйвер «{row.Driver}», порт «{row.Port}»)…");
            try
            {
                using var cls = new ManagementClass("Win32_Printer");
                using var np = cls.CreateInstance();
                np["DeviceID"] = clean;
                np["DriverName"] = row.Driver;
                np["PortName"] = row.Port;
                np.Put();
            }
            catch (Exception wmiEx)
            {
                _log.Warn("WMI-создание очереди не удалось: " + wmiEx.Message + ". Пробую Add-Printer…");
                var psi = new ProcessStartInfo("powershell.exe",
                    $"-NoProfile -NonInteractive -Command \"Add-Printer -Name '{clean.Replace("'", "''")}' " +
                    $"-DriverName '{row.Driver.Replace("'", "''")}' -PortName '{row.Port.Replace("'", "''")}'\"")
                { UseShellExecute = false, CreateNoWindow = true };
                using var proc = Process.Start(psi);
                proc!.WaitForExit(30000);
            }

            if (!ListPrinterRows().Any(r => r.Name.Equals(clean, StringComparison.OrdinalIgnoreCase)))
            {
                _log.Error($"Новая очередь «{clean}» не появилась — старую не трогаю.");
                return null;
            }

            if (row.Default) SetDefaultPrinter(clean);
            DeletePrinterQueue(oldName);
            _log.Info($"Очередь пересоздана начисто: «{clean}».");
            return clean;
        }
        catch (Exception ex)
        {
            _log.Error($"Пересоздание очереди «{oldName}»: {ex.Message}");
            return null;
        }
    }

    /// <summary>Версии файлов драйверов реальных принтеров — чтобы понимать, что именно установлено.</summary>
    private void LogDriverVersions()
    {
        try
        {
            var used = ListPrinterRows().Where(p => !IsVirtual(p.Name))
                .Select(p => p.Driver).Distinct(StringComparer.OrdinalIgnoreCase).ToHashSet(StringComparer.OrdinalIgnoreCase);
            if (used.Count == 0) return;

            using var drivers = new ManagementObjectSearcher("SELECT * FROM Win32_PrinterDriver");
            foreach (var d in drivers.Get().Cast<ManagementObject>())
            {
                var baseName = ((string)d["Name"]).Split(',')[0];
                if (!used.Contains(baseName)) continue;
                var parts = new List<string>();
                foreach (var key in new[] { "DriverPath", "ConfigFile", "DataFile" })
                {
                    if (d[key] is not string path || !File.Exists(path)) continue;
                    string ver;
                    try { ver = FileVersionInfo.GetVersionInfo(path).FileVersion ?? "?"; }
                    catch { ver = "?"; }
                    parts.Add($"{Path.GetFileName(path)} v{ver} ({File.GetLastWriteTime(path):yyyy-MM-dd})");
                }
                _log.Info($"Файлы драйвера «{baseName}»: {string.Join(", ", parts)}");
            }
        }
        catch (Exception ex) { _log.Warn("Версии драйверов: " + ex.Message); }
    }

    /// <summary>Какие очереди используют драйвер, в состав которого входит файл moduleFile.</summary>
    private List<string> QueuesUsingDriverFile(string moduleFile)
    {
        var result = new List<string>();
        if (string.IsNullOrWhiteSpace(moduleFile)) return result;
        try
        {
            var file = Path.GetFileName(moduleFile);
            var driverBases = new List<string>();
            using (var drivers = new ManagementObjectSearcher("SELECT * FROM Win32_PrinterDriver"))
            {
                foreach (var d in drivers.Get().Cast<ManagementObject>())
                {
                    var paths = new List<string?>
                    {
                        d["DriverPath"] as string, d["ConfigFile"] as string, d["DataFile"] as string
                    };
                    if (d["DependentFiles"] is string[] dep) paths.AddRange(dep);
                    if (paths.Any(path => path != null &&
                            Path.GetFileName(path).Equals(file, StringComparison.OrdinalIgnoreCase)))
                    {
                        var name = (string)d["Name"]; // "Имя,версия,окружение"
                        driverBases.Add(name.Split(',')[0]);
                    }
                }
            }
            if (driverBases.Count == 0) return result;
            result = ListPrinterRows()
                .Where(p => driverBases.Contains(p.Driver, StringComparer.OrdinalIgnoreCase))
                .Select(p => p.Name).ToList();
        }
        catch (Exception ex)
        {
            _log.Warn("Поиск драйвера по модулю: " + ex.Message);
        }
        return result;
    }

    private List<string> AllRealQueues()
    {
        try { return ListPrinterRows().Where(p => !IsVirtual(p.Name)).Select(p => p.Name).ToList(); }
        catch { return new List<string>(); }
    }

    /// <summary>Отключает/включает «расширенные функции печати» (EMF): RawOnly=true лечит
    /// падения приложений внутри GDI-драйверов.</summary>
    private void SetRawOnly(string printerName, bool on)
    {
        _log.Info($"«{printerName}»: RawOnly = {on} (расширенные функции печати {(on ? "выключены" : "включены")})…");
        using var searcher = new ManagementObjectSearcher(
            $"SELECT * FROM Win32_Printer WHERE Name = '{printerName.Replace("'", "''")}'");
        foreach (var printer in searcher.Get().Cast<ManagementObject>())
        {
            printer["RawOnly"] = on;
            printer.Put();
            _log.Info($"«{printerName}»: RawOnly применён.");
        }
        if (on) MarkTreatment(printerName, 1);
    }

    private static string? ReadLastPrinterFromPrefs(string prefsPath)
    {
        using var doc = JsonDocument.Parse(File.ReadAllText(prefsPath));
        if (!doc.RootElement.TryGetProperty("printing", out var printing) ||
            !printing.TryGetProperty("print_preview_sticky_settings", out var sticky) ||
            !sticky.TryGetProperty("appState", out var appStateEl))
            return null;

        var appState = appStateEl.GetString();
        if (string.IsNullOrEmpty(appState)) return null;

        using var app = JsonDocument.Parse(appState);
        if (app.RootElement.TryGetProperty("recentDestinations", out var dests) &&
            dests.ValueKind == JsonValueKind.Array && dests.GetArrayLength() > 0 &&
            dests[0].TryGetProperty("id", out var id))
            return id.GetString();
        return null;
    }

    // ---------- УПРАВЛЕНИЕ ПРИНТЕРАМИ (ручное удаление) ----------

    public sealed record PrinterOverview(string Name, string Driver, string Port, bool Default, bool Virtual);

    public List<PrinterOverview> GetPrinterOverview() =>
        ListPrinterRows()
            .Select(p => new PrinterOverview(p.Name, p.Driver, p.Port, p.Default, IsVirtual(p.Name)))
            .OrderBy(p => p.Virtual).ThenBy(p => p.Name)
            .ToList();

    /// <summary>Удаляет выбранные очереди, затем их осиротевшие драйверы и TCP/IP-порты.</summary>
    public string RemovePrinters(IReadOnlyCollection<string> names)
    {
        _log.Info("--- Удаление принтеров: " + string.Join("; ", names.Select(n => $"«{n}»")) + " ---");
        var before = ListPrinterRows();
        var selected = before.Where(r => names.Contains(r.Name)).ToList();

        var removedQueues = 0;
        foreach (var row in selected)
        {
            try { DeletePrinterQueue(row.Name); removedQueues++; }
            catch (Exception ex) { _log.Error($"Не удалось удалить очередь «{row.Name}»: {ex.Message}"); }
        }

        // Перезапуск спулера освобождает файлы драйверов
        try { RestartSpoolerAndClearQueue(); }
        catch (Exception ex) { _log.Warn("Перезапуск службы печати: " + ex.Message); }

        var remaining = ListPrinterRows();
        var removedDrivers = 0;
        foreach (var driver in selected.Select(s => s.Driver).Distinct(StringComparer.OrdinalIgnoreCase))
        {
            if (remaining.Any(r => r.Driver.Equals(driver, StringComparison.OrdinalIgnoreCase)))
            {
                _log.Info($"Драйвер «{driver}» ещё используется другими принтерами — не трогаю.");
                continue;
            }
            if (DeletePrinterDriver(driver)) removedDrivers++;
        }

        var removedPorts = 0;
        var usedPorts = remaining.Select(r => r.Port).ToHashSet(StringComparer.OrdinalIgnoreCase);
        foreach (var port in selected.Select(s => s.Port).Distinct(StringComparer.OrdinalIgnoreCase))
        {
            if (usedPorts.Contains(port)) continue;
            if (DeleteTcpIpPort(port)) removedPorts++;
        }

        var summary = $"Удалено: очередей — {removedQueues}, драйверов — {removedDrivers}, портов — {removedPorts}.";
        _log.Info(summary);
        return summary;
    }

    private bool DeletePrinterDriver(string driverBaseName)
    {
        try
        {
            using var searcher = new ManagementObjectSearcher("SELECT * FROM Win32_PrinterDriver");
            var ok = false;
            foreach (var d in searcher.Get().Cast<ManagementObject>())
            {
                var name = (string)d["Name"];
                if (!name.Split(',')[0].Equals(driverBaseName, StringComparison.OrdinalIgnoreCase)) continue;
                d.Delete();
                _log.Info($"Драйвер «{name}» удалён.");
                ok = true;
            }
            if (!ok) _log.Info($"Драйвер «{driverBaseName}» не найден (возможно, уже удалён).");
            return ok;
        }
        catch (Exception ex)
        {
            _log.Warn($"Не удалось удалить драйвер «{driverBaseName}»: {ex.Message}");
            return false;
        }
    }

    private bool DeleteTcpIpPort(string portName)
    {
        try
        {
            using var searcher = new ManagementObjectSearcher(
                $"SELECT * FROM Win32_TCPIPPrinterPort WHERE Name = '{portName.Replace("'", "''")}'");
            var ok = false;
            foreach (var port in searcher.Get().Cast<ManagementObject>())
            {
                port.Delete();
                _log.Info($"TCP/IP-порт «{portName}» удалён.");
                ok = true;
            }
            return ok;
        }
        catch (Exception ex)
        {
            _log.Warn($"Порт «{portName}» удалить не удалось (не TCP/IP-порт или занят): {ex.Message}");
            return false;
        }
    }

    private void DeletePrinterQueue(string printerName)
    {
        _log.Info($"Удаляю очередь печати «{printerName}»…");
        using var searcher = new ManagementObjectSearcher(
            $"SELECT * FROM Win32_Printer WHERE Name = '{printerName.Replace("'", "''")}'");
        foreach (var printer in searcher.Get().Cast<ManagementObject>())
        {
            printer.Delete();
            _log.Info($"Очередь «{printerName}» удалена.");
        }
    }

    private void SetDefaultPrinter(string printerName)
    {
        _log.Info($"Назначаю «{printerName}» принтером по умолчанию…");
        using var searcher = new ManagementObjectSearcher(
            $"SELECT * FROM Win32_Printer WHERE Name = '{printerName.Replace("'", "''")}'");
        foreach (var printer in searcher.Get().Cast<ManagementObject>())
        {
            var result = printer.InvokeMethod("SetDefaultPrinter", null);
            _log.Info($"SetDefaultPrinter(«{printerName}») → код {result}");
        }
    }

    private void CheckStuckJobs(List<CheckResult> results)
    {
        try
        {
            using var searcher = new ManagementObjectSearcher("SELECT * FROM Win32_PrintJob");
            var now = DateTime.Now;
            var stuckByPrinter = new Dictionary<string, int>();

            foreach (var job in searcher.Get().Cast<ManagementObject>())
            {
                var fullName = (string)(job["Name"] ?? "");          // "Принтер, 123"
                var printer = fullName.Split(',')[0].Trim();
                var status = (string)(job["JobStatus"] ?? "");
                DateTime submitted;
                try { submitted = ManagementDateTimeConverter.ToDateTime((string)job["TimeSubmitted"]); }
                catch { submitted = now; }

                var age = now - submitted;
                _log.Info($"Задание печати: «{fullName}», статус: «{status}», возраст: {age:hh\\:mm\\:ss}");

                if (age > TimeSpan.FromMinutes(10) ||
                    status.Contains("Error", StringComparison.OrdinalIgnoreCase))
                {
                    stuckByPrinter[printer] = stuckByPrinter.GetValueOrDefault(printer) + 1;
                }
            }

            foreach (var (printer, count) in stuckByPrinter)
                results.Add(new CheckResult
                {
                    Severity = Severity.Error,
                    Title = $"«{printer}»: застрявшие задания печати — {count} шт.",
                    Details = "Очередь будет очищена.",
                    Fix = () => { CancelAllJobs(printer); RestartSpoolerAndClearQueue(); }
                });

            if (stuckByPrinter.Count == 0)
                results.Add(new CheckResult { Severity = Severity.Ok, Title = "Застрявших заданий печати нет" });
        }
        catch (Exception ex)
        {
            _log.Error("Проверка заданий печати: " + ex.Message);
        }
    }

    private void CheckNetworkPorts(List<CheckResult> results)
    {
        try
        {
            using var printers = new ManagementObjectSearcher("SELECT Name, PortName FROM Win32_Printer");
            var usedPorts = printers.Get().Cast<ManagementObject>()
                .Where(p => !IsVirtual((string)p["Name"]))
                .Select(p => (Printer: (string)p["Name"], Port: (string)(p["PortName"] ?? "")))
                .ToList();

            using var searcher = new ManagementObjectSearcher("SELECT * FROM Win32_TCPIPPrinterPort");
            foreach (var port in searcher.Get().Cast<ManagementObject>())
            {
                var portName = (string)(port["Name"] ?? "");
                var owner = usedPorts.FirstOrDefault(u => u.Port == portName);
                if (owner.Printer is null) continue; // порт никем не используется

                var host = (string)(port["HostAddress"] ?? "");
                var tcpPort = Convert.ToInt32(port["PortNumber"] ?? 9100);
                _log.Info($"Проверка сетевого принтера «{owner.Printer}»: {host}:{tcpPort}…");

                if (TcpReachable(host, tcpPort, 1500))
                    results.Add(new CheckResult { Severity = Severity.Ok, Title = $"Сетевой принтер «{owner.Printer}» ({host}) доступен" });
                else
                    results.Add(new CheckResult
                    {
                        Severity = Severity.Error,
                        Title = $"Сетевой принтер «{owner.Printer}» недоступен по сети ({host}:{tcpPort})",
                        Details = "Проверьте: принтер включён, подключён к той же сети (Wi-Fi/кабель), " +
                                  "IP-адрес не поменялся. Если у принтера сменился IP — переустановите порт."
                    });
            }
        }
        catch (Exception ex)
        {
            _log.Error("Проверка сетевых портов: " + ex.Message);
        }
    }

    private static bool TcpReachable(string host, int port, int timeoutMs)
    {
        try
        {
            using var client = new TcpClient();
            return client.ConnectAsync(host, port).Wait(timeoutMs) && client.Connected;
        }
        catch { return false; }
    }

    private static bool IsVirtual(string name) =>
        VirtualPrinters.Any(v => name.Contains(v, StringComparison.OrdinalIgnoreCase));

    // ---------- ПОЧИНКА ----------

    /// <summary>Выполняет исправления. Возвращает число применённых фиксов.</summary>
    public int Fix(List<CheckResult> issues)
    {
        _log.Info("--- Починка ---");
        var fixes = issues.Where(i => i.IsProblem && i.Fix != null).ToList();
        var done = 0;
        var spoolerRestarted = false;

        foreach (var issue in fixes)
        {
            // Не перезапускаем спулер несколько раз подряд
            if (issue.Fix == (Action)RestartSpoolerAndClearQueue && spoolerRestarted) continue;
            try
            {
                _log.Info($"Исправляю: {issue.Title}");
                issue.Fix!();
                if (issue.Fix == (Action)RestartSpoolerAndClearQueue) spoolerRestarted = true;
                done++;
            }
            catch (Exception ex)
            {
                _log.Error($"Не удалось исправить «{issue.Title}»: {ex}");
            }
        }
        _log.Info($"Применено исправлений: {done} из {fixes.Count}");
        return done;
    }

    private void StartService(string name)
    {
        using var sc = new ServiceController(name);
        if (sc.Status == ServiceControllerStatus.Running) return;
        _log.Info($"Запускаю службу {name}…");
        sc.Start();
        sc.WaitForStatus(ServiceControllerStatus.Running, TimeSpan.FromSeconds(30));
        _log.Info($"Служба {name} запущена.");
    }

    public void RestartSpoolerAndClearQueue()
    {
        _log.Info("Перезапуск службы печати с очисткой очереди…");
        using var sc = new ServiceController("Spooler");

        if (sc.Status != ServiceControllerStatus.Stopped)
        {
            try
            {
                sc.Stop();
                sc.WaitForStatus(ServiceControllerStatus.Stopped, TimeSpan.FromSeconds(30));
            }
            catch (Exception ex)
            {
                _log.Warn("Служба не остановилась штатно: " + ex.Message + ". Завершаю процесс spoolsv.exe…");
                foreach (var proc in Process.GetProcessesByName("spoolsv"))
                {
                    try { proc.Kill(); proc.WaitForExit(5000); } catch { }
                }
            }
        }
        _log.Info("Служба печати остановлена.");

        // splwow64 — посредник печати 32/64: зависший экземпляр портит печать из браузеров
        foreach (var proc in Process.GetProcessesByName("splwow64"))
        {
            try { proc.Kill(); proc.WaitForExit(5000); _log.Info("Завершён зависший splwow64.exe."); }
            catch { }
        }

        // Очистка очереди
        var deleted = 0;
        if (Directory.Exists(SpoolDir))
        {
            foreach (var file in Directory.GetFiles(SpoolDir))
            {
                try { File.Delete(file); deleted++; }
                catch (Exception ex) { _log.Warn($"Не удалось удалить {Path.GetFileName(file)}: {ex.Message}"); }
            }
        }
        _log.Info($"Удалено файлов из очереди: {deleted}.");

        sc.Refresh();
        if (sc.Status != ServiceControllerStatus.Running)
        {
            sc.Start();
            sc.WaitForStatus(ServiceControllerStatus.Running, TimeSpan.FromSeconds(30));
        }
        _log.Info("Служба печати запущена заново.");
    }

    private void SetPrinterOnline(string printerName)
    {
        _log.Info($"Перевожу «{printerName}» в онлайн…");
        using var searcher = new ManagementObjectSearcher(
            $"SELECT * FROM Win32_Printer WHERE Name = '{printerName.Replace("'", "''")}'");
        foreach (var printer in searcher.Get().Cast<ManagementObject>())
        {
            printer["WorkOffline"] = false;
            printer.Put();
            _log.Info($"«{printerName}»: автономный режим выключен.");
        }
    }

    private void CancelAllJobs(string printerName)
    {
        _log.Info($"Отменяю все задания на «{printerName}»…");
        using var searcher = new ManagementObjectSearcher(
            $"SELECT * FROM Win32_Printer WHERE Name = '{printerName.Replace("'", "''")}'");
        foreach (var printer in searcher.Get().Cast<ManagementObject>())
        {
            var result = printer.InvokeMethod("CancelAllJobs", null);
            _log.Info($"CancelAllJobs(«{printerName}») → код {result}");
        }
    }

    private static string GetServiceStartMode(string name)
    {
        using var searcher = new ManagementObjectSearcher(
            $"SELECT StartMode FROM Win32_Service WHERE Name = '{name}'");
        foreach (var svc in searcher.Get().Cast<ManagementObject>())
            return (string)(svc["StartMode"] ?? "Unknown");
        return "Unknown";
    }

    private void SetServiceAutoStart(string name)
    {
        _log.Info($"Ставлю службе {name} автозапуск…");
        using var searcher = new ManagementObjectSearcher(
            $"SELECT * FROM Win32_Service WHERE Name = '{name}'");
        foreach (var svc in searcher.Get().Cast<ManagementObject>())
        {
            var inParams = svc.GetMethodParameters("ChangeStartMode");
            inParams["StartMode"] = "Automatic";
            var outParams = svc.InvokeMethod("ChangeStartMode", inParams, null);
            _log.Info($"ChangeStartMode({name}, Automatic) → код {outParams["ReturnValue"]}");
        }
        StartService(name);
    }
}
