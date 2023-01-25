using System;
using System.Data.SqlClient;
using System.IO;

namespace LabComm
{
    public class DBWatcher
    {
        static string FilePath = Directory.GetCurrentDirectory() + "/angry_bird.wav";
        System.Media.SoundPlayer player = new System.Media.SoundPlayer(FilePath);

        public void StartWatching()
        {
            SqlDependency.Stop(Processing.sqlConnectionStr);
            SqlDependency.Start(Processing.sqlConnectionStr);
            ExecuteWatchingQuery();
        }

        private void ExecuteWatchingQuery()
        {
            using (SqlConnection connection = new SqlConnection(Processing.sqlConnectionStr))
            {
                connection.Open();
                
                using (var command = new SqlCommand(
                    @"SET ANSI_NULLS ON
                    SET ANSI_PADDING ON
                    SET ANSI_WARNINGS ON
                    SET CONCAT_NULL_YIELDS_NULL ON
                    SET QUOTED_IDENTIFIER ON
                    SET NUMERIC_ROUNDABORT OFF
                    SET ARITHABORT ON
                    select patdirec_id, date_bio from dbo.patdirec", connection))
                {
                    var sqlDependency = new SqlDependency(command);
                    sqlDependency.OnChange += new OnChangeEventHandler(OnDatabaseChange);
                    command.ExecuteReader();
                }
            }
        }

        private void OnDatabaseChange(object sender, SqlNotificationEventArgs args)
        {
            SqlNotificationInfo info = args.Info;
            if (SqlNotificationInfo.Insert.Equals(info)
                || SqlNotificationInfo.Update.Equals(info)
                || SqlNotificationInfo.Delete.Equals(info))
            {
                //player.Play();
                Console.WriteLine("Забор биоматериала произведен!");
            }
            ExecuteWatchingQuery();
        }
    }
}
