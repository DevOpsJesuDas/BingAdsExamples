using System;
using System.Configuration;
using System.Data.SqlClient;
using System.IO;

namespace ReadingExcelFiles
{
	class Program
	{
		static void Main(string[] args)
		{
			WriteLog("=================Scheduler Started=========================================" + DateTime.Now);
			try
			{
				var connectionString = ConfigurationManager.ConnectionStrings["AMS"].ConnectionString;

				SqlConnection conn = new SqlConnection(connectionString);
				if (conn.State.ToString() != "1")
					conn.Open();

				WriteLog("Connection State:: " + conn.State.ToString());
				string sql = "";
				//Reading Bing Files
				WriteLog("Bing Read Files Starts");
				string path = System.Configuration.ConfigurationManager.AppSettings["BingFolderPath"].ToString(); // @"F:\AdWords";
				foreach (string fileName in Directory.GetFiles(path))
				{
					string[] read;
					char[] seperators = { ',' };

					StreamReader sr = new StreamReader(fileName);

					string data = sr.ReadLine();

					while ((data = sr.ReadLine()) != null)
					{
						if (data != "")
						{

							read = data.Split(seperators, StringSplitOptions.RemoveEmptyEntries);

							if (read[0].ToString().ToLower().Contains("timeperiod"))
							{
								data = sr.ReadLine();
								read = data.Split(seperators, StringSplitOptions.RemoveEmptyEntries);
								sql = "insert into BingAdWordsSpendByBranch (ID,Total,ClientID,SpendDate,DateCreated) values((select isnull(max(0),0)+1 from BingAdWordsSpendByBranch),@Total,@ClientID,CONVERT(VARCHAR(10),GETDATE()-1,111),getdate())";

								SqlCommand cmd = new SqlCommand(sql, conn);

								cmd.Parameters.Add("@Total", Convert.ToDecimal(read[3].Split('"')[1]));
								cmd.Parameters.Add("@ClientID", Convert.ToString(read[2].Split('"')[1]));
								cmd.ExecuteNonQuery();
								cmd.Dispose();
							}
						}
					}
					sr.Close();
					sr.Dispose();

					if ((System.IO.File.Exists(fileName)) && (System.Configuration.ConfigurationManager.AppSettings["DeleteFiles"].ToString() == "Yes"))
					{
						System.IO.File.Delete(fileName);
					}
					
				}
				WriteLog("Bing Read Files Ends");
				//Ends


			}
			catch (Exception ex)
			{
				WriteLog("************************************ExCeption Starts************************************" + DateTime.Now);
				WriteLog(ex.ToString());
				WriteLog("************************************ExCeption Ends**************************************");
				throw;
			}
			finally
			{
				WriteLog("=================Scheduler Ended=========================================" + DateTime.Now);
			}


		}

		public static void WriteLog(string strLog)
		{
			StreamWriter log;
			FileStream fileStream = null;
			DirectoryInfo logDirInfo = null;
			FileInfo logFileInfo;

			string logFilePath = System.Configuration.ConfigurationManager.AppSettings["LogFolderPath"].ToString();
			logFilePath = logFilePath + "Log_Bing-FileRead_" + System.DateTime.Today.ToString("MM-dd-yyyy") + "." + "txt";
			logFileInfo = new FileInfo(logFilePath);
			logDirInfo = new DirectoryInfo(logFileInfo.DirectoryName);
			if (!logDirInfo.Exists) logDirInfo.Create();
			if (!logFileInfo.Exists)
			{
				fileStream = logFileInfo.Create();
			}
			else
			{
				fileStream = new FileStream(logFilePath, FileMode.Append);
			}
			log = new StreamWriter(fileStream);
			log.WriteLine(strLog);
			log.Close();
			fileStream.Close();
		}

	}
}
