using System;

namespace AutoCAD_PIK_Manager
{
   public class LogAddin
   {
      private string _plugin = string.Empty;

      public LogAddin(string plugin)
      {
         _plugin = "Plugin " + plugin;            
      }

      /// <summary>
      /// Debug: сообщения отладки, профилирования. В production системе обычно сообщения этого уровня включаются при первоначальном запуске системы или для поиска узких мест (bottleneck-ов).
      /// </summary>
      /// <param name="message"></param>
      public void Debug(string message)
      {
         Log.Debug(_plugin + message);
      }

      public void Debug(string message, params object[] args)
      {
         Log.Debug(_plugin + message, args);
      }

      public void Debug(Exception ex, string message, params object[] args)
      {
         Log.Debug(ex, _plugin + message, args);
      }

      /// <summary>
      /// Error: ошибка в работе системы, требующая вмешательства. Что-то не сохранилось, что-то отвалилось. Необходимо принимать меры довольно быстро! Ошибки этого уровня и выше требуют немедленной записи в лог, чтобы ускорить реакцию на них.Нужно понимать, что ошибка пользователя – это не ошибка системы. Если пользователь ввёл в поле -1, где это не предполагалось – не надо писать об этом в лог ошибок.
      /// </summary>
      /// <param name="message"></param>
      public void Error(string message)
      {
         Log.Error(_plugin + message);
      }

      public void Error(Exception ex, string message, params object[] args)
      {
         Log.Error(ex, _plugin + message, args);
      }

      public void Error(string message, params object[] args)
      {
         Log.Error(_plugin + message, args);
      }

      /// <summary>
      /// Fatal: это особый класс ошибок. Такие ошибки приводят к неработоспособности системы в целом, или неработоспособности одной из подсистем.Чаще всего случаются фатальные ошибки из-за неверной конфигурации или отказов оборудования. Требуют срочной, немедленной реакции. Возможно, следует предусмотреть уведомление о таких ошибках по SMS.
      /// </summary>
      /// <param name="message"></param>
      public void Fatal(string message)
      {
         Log.Fatal(_plugin + message);
      }

      public void Fatal(Exception ex, string message, params object[] args)
      {
         Log.Fatal(ex, _plugin + message, args);
      }

      public void Fatal(string message, params object[] args)
      {
         Log.Fatal(_plugin + message, args);
      }

      /// <summary>
      /// Info: обычные сообщения, информирующие о действиях системы.Реагировать на такие сообщения вообще не надо, но они могут помочь, например, при поиске багов, расследовании интересных ситуаций итд.
      /// </summary>
      /// <param name="message"></param>
      public void Info(string message)
      {
         Log.Info(_plugin + message);
      }

      public void Info(string message, params object[] args)
      {
         Log.Info(_plugin + message, args);
      }

      public void Info(Exception ex, string message, params object[] args)
      {
         Log.Info(ex, _plugin + message, args);
      }

      /// <summary>
      /// Warn: записывая такое сообщение, система пытается привлечь внимание обслуживающего персонала.Произошло что-то странное. Возможно, это новый тип ситуации, ещё не известный системе. Следует разобраться в том, что произошло, что это означает, и отнести ситуацию либо к инфо-сообщению, либо к ошибке.Соответственно, придётся доработать код обработки таких ситуаций.
      /// </summary>
      /// <param name="message"></param>
      public void Warn(string message)
      {
         Log.Warn(_plugin + message);
      }

      public void Warn(string message, params object[] args)
      {
         Log.Warn(_plugin + message, args);
      }

      public void Warn(Exception ex, string message, params object[] args)
      {
         Log.Warn(ex, _plugin + message, args);
      }

      public void StartCommand(string message)
      {
         Log.Info(_plugin + " Start command: " + message);
      }
   }
}