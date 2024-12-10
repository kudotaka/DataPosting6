﻿﻿using System.Collections.Immutable;
using System.Text;
using ClosedXML.Excel;
using ClosedXML.Graphics;
using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Logging;
using Microsoft.Extensions.Options;
using Utf8StringInterpolation;
using ZLogger;
using ZLogger.Providers;

//==
var builder = ConsoleApp.CreateBuilder(args);
builder.ConfigureServices((ctx,services) =>
{
    // Register appconfig.json to IOption<MyConfig>
    services.Configure<MyConfig>(ctx.Configuration);

    // Using Cysharp/ZLogger for logging to file
    services.AddLogging(logging =>
    {
        logging.ClearProviders();
        logging.SetMinimumLevel(LogLevel.Trace);
        var jstTimeZoneInfo = TimeZoneInfo.FindSystemTimeZoneById("Tokyo Standard Time");
        var utcTimeZoneInfo = TimeZoneInfo.Utc;
        logging.AddZLoggerConsole(options =>
        {
            options.UsePlainTextFormatter(formatter => 
            {
                formatter.SetPrefixFormatter($"{0:yyyy-MM-dd'T'HH:mm:sszzz}|{1:short}|", (in MessageTemplate template, in LogInfo info) => template.Format(TimeZoneInfo.ConvertTime(info.Timestamp.Utc, jstTimeZoneInfo), info.LogLevel));
                formatter.SetExceptionFormatter((writer, ex) => Utf8String.Format(writer, $"{ex.Message}"));
            });
        });
        logging.AddZLoggerRollingFile(options =>
        {
            options.UsePlainTextFormatter(formatter => 
            {
                formatter.SetPrefixFormatter($"{0:yyyy-MM-dd'T'HH:mm:sszzz}|{1:short}|", (in MessageTemplate template, in LogInfo info) => template.Format(TimeZoneInfo.ConvertTime(info.Timestamp.Utc, jstTimeZoneInfo), info.LogLevel));
                formatter.SetExceptionFormatter((writer, ex) => Utf8String.Format(writer, $"{ex.Message}"));
            });

            // File name determined by parameters to be rotated
            options.FilePathSelector = (timestamp, sequenceNumber) => $"logs/{timestamp.ToLocalTime():yyyy-MM-dd}_{sequenceNumber:00}.log";
            
            // The period of time for which you want to rotate files at time intervals.
            options.RollingInterval = RollingInterval.Day;
            
            // Limit of size if you want to rotate by file size. (KB)
            options.RollingSizeKB = 1024;        
        });
    });
});

var app = builder.Build();
app.AddCommands<DataPortingApp>();
app.Run();


public class DataPortingApp : ConsoleAppBase
{
    bool isAllPass = true;

    readonly ILogger<DataPortingApp> logger;
    readonly IOptions<MyConfig> config;

    Dictionary<MyTypes, List<MyDefinition>> dicMyDefinition = new Dictionary<MyTypes, List<MyDefinition>>();

    public DataPortingApp(ILogger<DataPortingApp> logger,IOptions<MyConfig> config)
    {
        this.logger = logger;
        this.config = config;
    }

//    [Command("")]
    public void Porting(string orignal, string format, string outpath)
    {
//== start
        logger.ZLogInformation($"==== tool {getMyFileVersion()} ====");
        if (!File.Exists(orignal))
        {
            logger.ZLogError($"[NG] エクセルファイルが見つかりません{orignal}");
            return;
        }
        if (!File.Exists(format))
        {
            logger.ZLogError($"[NG] エクセルファイルが見つかりません{format}");
            return;
        }

        string orignalTypeCell = config.Value.OrignalTypeCell;
        string formatSheetName = config.Value.FormatSheetName;

        readDefinition(dicMyDefinition);
        printDefinition(dicMyDefinition);

        portingExcel(orignal, orignalTypeCell, format, outpath, dicMyDefinition);


//== finish
        if (isAllPass)
        {
            logger.ZLogInformation($"== [Congratulations!] すべての処理をパスしました ==");
        }
        logger.ZLogInformation($"==== tool finish ====");
    }

    private void readDefinition(Dictionary<MyTypes, List<MyDefinition>> dic)
    {
        string definition6UOrignalToFormat = config.Value.Definition6UOrignalToFormat;
        List<MyDefinition> list6U = new List<MyDefinition>();
        foreach (var keyAndValue in definition6UOrignalToFormat.Split(','))
        {
            string[] item = keyAndValue.Split('|');
            var definition = new MyDefinition();
            definition.orignalCell = item[0];
            definition.portingCell = item[1];
            list6U.Add(definition);
        }
        dic.Add(MyTypes.Type6U, list6U);

        string definition14UOrignalToFormat = config.Value.Definition14UOrignalToFormat;
        List<MyDefinition> list14U = new List<MyDefinition>();
        foreach (var keyAndValue in definition14UOrignalToFormat.Split(','))
        {
            string[] item = keyAndValue.Split('|');
            var definition = new MyDefinition();
            definition.orignalCell = item[0];
            definition.portingCell = item[1];
            list14U.Add(definition);
        }
        dic.Add(MyTypes.Type14U, list14U);
    }

    private void printDefinition(Dictionary<MyTypes, List<MyDefinition>> dic)
    {
        logger.ZLogTrace($"== start print ==");
        foreach (var key in dic.Keys)
        {
            foreach (var definition in dic[key])
            {
                logger.ZLogTrace($"キー:{convertTypesToReadableTypes(key)} {definition.orignalCell}-->{definition.portingCell}");
            }
        }
        logger.ZLogTrace($"== end print ==");
    }

    private string convertTypesToReadableTypes(MyTypes types)
    {
        switch (types)
        {
            case MyTypes.Type6U:
                return "6U";
            case MyTypes.Type14U:
                return "14U";
            case MyTypes.UnKnown:
                return "不明";
            default:
                break;
        }
        return types.ToString();
    }

    private void portingExcel(string excel, string orignalTypeCell, string format, string outpath, Dictionary<MyTypes, List<MyDefinition>> dic)
    {
        logger.ZLogInformation($"== start portingExcel ==");
        bool isError = false;
        try
        {
            File.Copy(format, outpath, true);
        }
        catch (System.Exception)
        {
            
            throw;
        }
        using FileStream fsFormatExcel = new FileStream(outpath, FileMode.Open, FileAccess.ReadWrite, FileShare.Write);
        using XLWorkbook xlWorkbookFormatExcel = new XLWorkbook(fsFormatExcel);

        using FileStream fsOrignalExcel = new FileStream(excel, FileMode.Open, FileAccess.Read, FileShare.Read);
        using XLWorkbook xlWorkbookOrignalExcel = new XLWorkbook(fsOrignalExcel);
        IXLWorksheets sheetsOrignalExcel = xlWorkbookOrignalExcel.Worksheets;
        foreach (IXLWorksheet? sheetOrignal in sheetsOrignalExcel)
        {
            IXLCell cellColumn = sheetOrignal.Cell(orignalTypeCell);
            logger.ZLogDebug($"cell is Text type. value:{cellColumn.GetValue<string>()}");
            var type = convertWordToType(cellColumn.GetValue<string>());
            switch (type)
            {
                case MyTypes.Type6U:
                    logger.ZLogDebug($"6U! {sheetOrignal.Name},{covertTypeToSheetName(type)}");
                    break;
                case MyTypes.Type14U:
                    logger.ZLogDebug($"14U! {sheetOrignal.Name},{covertTypeToSheetName(type)}");
                    break;                
                default:
                    break;
            }

            foreach (IXLWorksheet? sheetFormat in xlWorkbookFormatExcel.Worksheets)
            {
                    logger.ZLogDebug($"sheetFormat {sheetFormat.Name}");
            }

            // copy sheet
            var formatSheet = xlWorkbookFormatExcel.Worksheet(covertTypeToSheetName(type));
            var portingSheet = formatSheet.CopyTo(sheetOrignal.Name);
            // copy cell
            foreach (var definition in dic[type])
            {
                portingSheet.Cell(definition.portingCell).Value = sheetOrignal.Cell(definition.orignalCell).Value;
            }
        }

        // detele formatSheet
        xlWorkbookFormatExcel.Worksheet(covertTypeToSheetName(MyTypes.Type6U)).Delete();
        xlWorkbookFormatExcel.Worksheet(covertTypeToSheetName(MyTypes.Type14U)).Delete();

        // save
        xlWorkbookFormatExcel.Save();


        if (!isError)
        {
            logger.ZLogInformation($"[OK] readOrignalExcel()は正常に処理できました");
        }
        else
        {
            isAllPass = false;
            logger.ZLogError($"[NG] readOrignalExcel()でエラーが発生しました");
        }
        logger.ZLogInformation($"== end portingExcel ==");
    }

    private MyTypes convertWordToType(string word)
    {
        Dictionary<string, MyTypes> dic = new Dictionary<string, MyTypes>();
        string definition14UOrignalToFormat = config.Value.DefinitionWordToType;
        foreach (var keyAndValue in definition14UOrignalToFormat.Split(','))
        {
            string[] item = keyAndValue.Split('|');
            dic.Add(item[0], (MyTypes)int.Parse(item[1]));
        }
        if (dic.ContainsKey(word))
        {
            return dic[word];
        }

        return MyTypes.UnKnown;
    }

    private string covertTypeToSheetName(MyTypes type)
    {
        Dictionary<MyTypes, string> dic = new Dictionary<MyTypes, string>();
        string formatSheetName = config.Value.FormatSheetName;
        foreach (var keyAndValue in formatSheetName.Split(','))
        {
            string[] item = keyAndValue.Split('|');
            dic.Add((MyTypes)int.Parse(item[0]), item[1]);
        }
        if (dic.ContainsKey(type))
        {
            return dic[type];
        }

        return "";
    }



    private string getMyFileVersion()
    {
        System.Diagnostics.FileVersionInfo ver = System.Diagnostics.FileVersionInfo.GetVersionInfo(System.Reflection.Assembly.GetExecutingAssembly().Location);
        return ver.InternalName + "(" + ver.FileVersion + ")";
    }
}

//==
public class MyConfig
{
    public string Definition6UOrignalToFormat {get; set;} = "";
    public string Definition14UOrignalToFormat {get; set;} = "";
    public string DefinitionWordToType {get; set;} = "";
    public string OrignalTypeCell {get; set;} = "";
    public string FormatSheetName {get; set;} = "";
}

public enum MyTypes
{
    Type6U = 6,
    Type14U = 14,
    UnKnown = 91
}

public class MyDefinition
{
    public string orignalCell { set; get; } = "";
    public string portingCell { set; get; } = "";
}