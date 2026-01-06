<?xml version="1.0" encoding="utf-8" ?>
<log4net>
  <appender name="MasterFileAppender" type="log4net.Appender.FileAppender">
    <file type="log4net.Util.PatternString" value="Logs\Master_Execution_%date{yyyyMMdd_HHmm}.log" />
    <appendToFile value="true" />
    <layout type="log4net.Layout.PatternLayout">
      <conversionPattern value="%date [%thread] %-5level %logger - %message%newline" />
    </layout>
  </appender>
  <root>
    <level value="ALL" />
    <appender-ref ref="MasterFileAppender" />
  </root>
</log4net>


log4net.Config.XmlConfigurator.Configure(new System.IO.FileInfo("log4net.config"));


protected ILog log = LogManager.GetLogger(typeof(BaseTest));
    private FileAppender _individualTestAppender;

    [SetUp]
    public void StartTestLogging()
    {
        string testName = TestContext.CurrentContext.Test.Name;
        // Path for the individual test log
        string logPath = $"Logs\\IndividualTests\\{testName}.log";

        // Create a new appender for THIS test only
        _individualTestAppender = new FileAppender
        {
            Name = testName,
            File = logPath,
            AppendToFile = false, // Fresh file for each test
            Layout = new log4net.Layout.PatternLayout("%date %-5level - %message%newline")
        };
        _individualTestAppender.ActivateOptions();

        // Attach this appender to the log4net hierarchy
        var hierarchy = (Hierarchy)LogManager.GetRepository();
        hierarchy.Root.AddAppender(_individualTestAppender);
        hierarchy.RaiseConfigurationChanged(System.EventArgs.Empty);

        log.Info($"--- START OF TEST: {testName} ---");
    }



public void EndTestLogging()
    {
        log.Info($"--- END OF TEST: {TestContext.CurrentContext.Test.Name} ---");

        // Remove the individual appender so the next test doesn't write to this file
        var hierarchy = (Hierarchy)LogManager.GetRepository();
        hierarchy.Root.RemoveAppender(_individualTestAppender);
        _individualTestAppender.Close();
    }
