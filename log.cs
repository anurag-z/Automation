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
