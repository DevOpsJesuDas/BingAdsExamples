using System;
using System.Linq;
using System.Configuration;
using System.Net.Http;
using System.ServiceModel;
using System.Threading.Tasks;
using Microsoft.BingAds;
using Microsoft.BingAds.V12.CustomerManagement;
using BingAdsConsoleApp.Properties;
using BingAdsExamplesLibrary;
using System.IO;
using System.Net.Mail;

namespace BingAdsConsoleApp
{
	class Program
	{
		/// <summary>
		/// Log File Name
		/// </summary>
		public static string logFineName = null;

		/// <summary>
		/// Error 
		/// </summary>
		public static string errorExisted = null;

		// Uncomment any examples that you want to run. 
		private static readonly ExampleBase[] _examples =
		{

			//new BingAdsExamplesLibrary.V12.Labels(),
			//new BingAdsExamplesLibrary.V12.OfflineConversions(),
			//new BingAdsExamplesLibrary.V12.KeywordPlanner(),
			//new BingAdsExamplesLibrary.V12.BudgetOpportunities(),
			//new BingAdsExamplesLibrary.V12.BulkServiceManagerDemo(),
			//new BingAdsExamplesLibrary.V12.BulkAdExtensions(),
			//new BingAdsExamplesLibrary.V12.AdExtensions(),
			//new BingAdsExamplesLibrary.V12.BulkKeywordsAds(),
			//new BingAdsExamplesLibrary.V12.KeywordsAds(),
			//new BingAdsExamplesLibrary.V12.BulkNegativeKeywords(),
			//new BingAdsExamplesLibrary.V12.NegativeKeywords(),
			//new BingAdsExamplesLibrary.V12.BulkAdGroupUpdate(),
			//new BingAdsExamplesLibrary.V12.BulkProductPartitionUpdateBid(),
			//new BingAdsExamplesLibrary.V12.ConversionGoals(),
			//new BingAdsExamplesLibrary.V12.BulkRemarketingLists(),
			//new BingAdsExamplesLibrary.V12.RemarketingLists(),
			//new BingAdsExamplesLibrary.V12.BulkShoppingCampaigns(),
			//new BingAdsExamplesLibrary.V12.ShoppingCampaigns(),
			//new BingAdsExamplesLibrary.V12.DynamicSearchCampaigns(),
			//new BingAdsExamplesLibrary.V12.BulkTargetCriterions(),
			//new BingAdsExamplesLibrary.V12.TargetCriterions(),
			//new BingAdsExamplesLibrary.V12.GeographicalLocations(),
			//new BingAdsExamplesLibrary.V12.BulkNegativeSites(),
			//new BingAdsExamplesLibrary.V12.InviteUser(),
			//new BingAdsExamplesLibrary.V12.CustomerSignup(),
			//new BingAdsExamplesLibrary.V12.ManageClient(),
			new BingAdsExamplesLibrary.V12.ReportRequests(),

			//new BingAdsExamplesLibrary.V12.SearchUserAccounts(),
		};

		private static AuthorizationData _authorizationData;
		private static string ClientState = "FL";
		private static ServiceClient<ICustomerManagementService> _customerManagementService;

		static void Main(string[] args)
		{
			WriteLog("=============================Scheduler Started=========================================" + DateTime.Now);
			try
			{
				WriteLog("Auth Started");
				Authentication authentication = AuthenticateWithOAuth();
				WriteLog("Auth Ended");
				// Most Bing Ads service operations require account and customer ID. 
				// This utiltiy operation sets the global authorization data instance 
				// to the first account that the current authenticated user can access. 
				SetAuthorizationDataAsync(authentication).Wait();


				// Run all of the examples that were included above.
				//foreach (var example in _examples)
				//{
				//	example.RunAsync(_authorizationData).Wait();

				//}


			}
			// Catch authentication exceptions
			catch (OAuthTokenRequestException ex)
			{
				OutputStatusMessage(string.Format("OAuthTokenRequestException Message:\n{0}", ex.Message));
				if (ex.Details != null)
				{
					OutputStatusMessage(string.Format("OAuthTokenRequestException Details:\nError: {0}\nDescription: {1}",
					ex.Details.Error, ex.Details.Description));
				}
				WriteLog("**********************************OAuthTokenRequestException Starts**********************************");
				WriteLog(ex.Details.Description);
				WriteLog("**********************************OAuthTokenRequestException Ends**********************************");
				errorExisted = "yes";
			}
			// Catch Customer Management service exceptions
			catch (FaultException<AdApiFaultDetail> ex)
			{
				OutputStatusMessage(string.Join("; ", ex.Detail.Errors.Select(error =>
				{
					if ((error.Code == 105) || (error.Code == 106))
					{
						return "Authorization data is missing or incomplete for the specified environment.\n" +
							   "To run the examples switch users or contact support for help with the following error.\n";
					}
					WriteLog("*********************" + error.Code + "::" + error.Message + "*********************");

					return string.Format("{0}: {1}", error.Code, error.Message);
				})));

				OutputStatusMessage(string.Join("; ",
					ex.Detail.Errors.Select(error => string.Format("{0}: {1}", error.Code, error.Message))));

				WriteLog("*********************FaultException*********************");
				errorExisted = "yes";

			}
			catch (FaultException<Microsoft.BingAds.V12.CustomerManagement.ApiFault> ex)
			{
				OutputStatusMessage(string.Join("; ",
					ex.Detail.OperationErrors.Select(error => string.Format("{0}: {1}", error.Code, error.Message))));
				WriteLog("*********************FaultException*********************");
				WriteLog("" + ex.Message + "");
				errorExisted = "yes";
			}
			catch (HttpRequestException ex)
			{
				WriteLog("*********************HttpRequestException*********************");
				WriteLog("" + ex.Message + "");
				OutputStatusMessage(ex.Message);
			}
			finally
			{
				WriteLog("==============================Scheduler Ended=========================================" + DateTime.Now);

				System.Net.Mail.MailMessage msg;
				System.Net.Mail.SmtpClient smtpServer = null;
				System.Net.NetworkCredential credentials = null;

				msg = new System.Net.Mail.MailMessage();
				msg.Body = "Hi Team, <br><br> Please find the attached file for the Bing AdWords log details. <br> <br> Thanks.";
				msg.From = new MailAddress("amsadmin@allmysons.com");
				//msg.To.Add(new MailAddress("ram@allmysons.com"));
				//msg.To.Add(new MailAddress("pbajwa@allmysons.com"));
				//msg.To.Add(new MailAddress("ashraf.syed@allmysons.com"));
				msg.To.Add(new MailAddress("jesu.danasiri@gmail.com"));

				if (errorExisted == "yes")
				{
					msg.Subject = "Bing AdWords Scheduler Status on " + System.DateTime.Today.ToString("MM-dd-yyyy") + " ****** Error Occurred **********";
				}
				else
				{
					msg.Subject = "Bing AdWords Scheduler Status on " + System.DateTime.Today.ToString("MM-dd-yyyy") + " ******* Run Successfully ******";
				}

				smtpServer = new System.Net.Mail.SmtpClient("mail.allmysons.com");
				credentials = new System.Net.NetworkCredential("amsadmin@allmysons.com", "ams2012");

				MemoryStream ms1 = new MemoryStream(File.ReadAllBytes(logFineName));
				Attachment att1 = new System.Net.Mail.Attachment(ms1, "LogFile.txt");
				msg.Attachments.Add(att1);

				msg.IsBodyHtml = true;
				smtpServer.Credentials = credentials;
				smtpServer.Send(msg);

			}
		}

		private static Authentication AuthenticateWithOAuth()
		{
			var apiEnvironment =
				ConfigurationManager.AppSettings["BingAdsEnvironment"] = ApiEnvironment.Production.ToString();
			var oAuthDesktopMobileAuthCodeGrant = new OAuthDesktopMobileAuthCodeGrant(
				Settings.Default["ClientId"].ToString(),
				apiEnvironment);
			try
			{

				// It is recommended that you specify a non guessable 'state' request parameter to help prevent
				// cross site request forgery (CSRF). 
				oAuthDesktopMobileAuthCodeGrant.State = ClientState;

				string refreshToken;

				// If you have previously securely stored a refresh token, try to use it.
				if (GetRefreshToken(out refreshToken))
				{
					AuthorizeWithRefreshTokenAsync(oAuthDesktopMobileAuthCodeGrant, refreshToken).Wait();
				}
				else
				{
					oAuthDesktopMobileAuthCodeGrant.GetAuthorizationEndpoint();
					// You must request user consent at least once through a web browser control. 
					// Call the GetAuthorizationEndpoint method of the OAuthDesktopMobileAuthCodeGrant instance that you created above.
					Console.WriteLine(string.Format(
						"The Bing Ads user must provide consent for your application to access their Bing Ads accounts.\n" +
						"Open a new web browser and navigate to {0}.\n\n" +
						"After the user has granted consent in the web browser for the application to access " +
						"their Bing Ads accounts, please enter the response URI that includes " +
						"the authorization 'code' parameter: \n", oAuthDesktopMobileAuthCodeGrant.GetAuthorizationEndpoint()));

					// Request access and refresh tokens using the URI that you provided manually during program execution.
					//var responseUri = new Uri("https://login.live.com/oauth20_desktop.srf?code=M29529e7a-79aa-818e-95f8-b358c406f67f&state=FL&lc=1033");
					var responseUri = new Uri("https://login.live.com/oauth20_desktop.srf?code=Medec11b4-6009-e24f-b8f9-d53a0a096749&state=FL.&lc=1033");
					if (oAuthDesktopMobileAuthCodeGrant.State != ClientState)
					{
						WriteLog("The OAuth response state does not match the client request state");
						throw new HttpRequestException("The OAuth response state does not match the client request state.");
					}
					WriteLog("RequestAccessAndRefreshTokensAsync== Starts");
					oAuthDesktopMobileAuthCodeGrant.RequestAccessAndRefreshTokensAsync(responseUri).Wait();
					WriteLog("RequestAccessAndRefreshTokensAsync== Ends");
					SaveRefreshToken(oAuthDesktopMobileAuthCodeGrant.OAuthTokens.RefreshToken);
				}

				// It is important to save the most recent refresh token whenever new OAuth tokens are received. 
				// You will want to subscribe to the NewOAuthTokensReceived event handler. 
				// When calling Bing Ads services with ServiceClient<TService>, BulkServiceManager, or ReportingServiceManager, 
				// each instance will refresh your access token automatically if they detect the AuthenticationTokenExpired (109) error code. 
				oAuthDesktopMobileAuthCodeGrant.NewOAuthTokensReceived +=
						(sender, tokens) => SaveRefreshToken(tokens.NewRefreshToken);
				
			}
			catch (Exception ex)
			{
				errorExisted = "yes";
				WriteLog("***********************************AuthenticateWithOAuth Exception==" + ex.ToString());
			}
			return oAuthDesktopMobileAuthCodeGrant;
		}

		/// <summary>
		/// Write to the console by default.
		/// </summary>
		/// <param name="msg">The message sent to console output.</param>
		private static void OutputStatusMessage(String msg)
		{
			Console.WriteLine(msg);
		}

		private static bool GetRefreshToken(out string refreshToken)
		{
			var filePath = System.Configuration.ConfigurationManager.AppSettings["RefreshTokenPath"].ToString() + @"\refreshtoken.txt";
			if (!File.Exists(filePath))
			{
				refreshToken = null;
				return false;
			}

			String fileContents;
			using (StreamReader sr = new StreamReader(filePath))
			{
				fileContents = sr.ReadToEnd();
			}
			if (string.IsNullOrEmpty(fileContents))
			{
				refreshToken = null;
				return false;
			}

			try
			{
				WriteLog("================= GetRefreshToken Success");
				refreshToken = fileContents;
				return true;
			}
			catch (FormatException ex)
			{
				errorExisted = "yes";
				WriteLog("****************************GetRefreshToken FormatException==>" + ex);
				refreshToken = null;
				return false;
			}
		}

		private static Task<OAuthTokens> AuthorizeWithRefreshTokenAsync(OAuthDesktopMobileAuthCodeGrant authentication, string refreshToken)
		{
			return authentication.RequestAccessAndRefreshTokensAsync(refreshToken);
		}

		private static void SaveRefreshToken(string newRefreshtoken)
		{
			if (newRefreshtoken != null)
			{
				using (StreamWriter outputFile = new StreamWriter(
				System.Configuration.ConfigurationManager.AppSettings["LogFolderPath"].ToString() + @"\refreshtoken.txt",
				false))
				{
					outputFile.WriteLine(newRefreshtoken);
					WriteLog("=================New Reresh Token ::" + newRefreshtoken + "::=========================================");
				}
			}
		}

		/// <summary>
		/// Utility method for setting the customer and account identifiers within the global 
		/// <see cref="_authorizationData"/> instance. 
		/// </summary>
		/// <param name="authentication">The OAuth authentication credentials.</param>
		/// <returns></returns>
		private static async Task SetAuthorizationDataAsync(Authentication authentication)
		{
			try
			{
				_authorizationData = new AuthorizationData
				{
					Authentication = authentication,
					DeveloperToken = Settings.Default["DeveloperToken"].ToString()
					//AccountId = 375726,
					//CustomerId= 11160683
				};

				_customerManagementService = new ServiceClient<ICustomerManagementService>(_authorizationData);

				var getUserRequest = new GetUserRequest
				{
					UserId = null
				};
				var getUserResponse = (await _customerManagementService.CallAsync((s, r) => s.GetUserAsync(r), getUserRequest));
				var user = getUserResponse.User;

				var predicate = new Predicate
				{
					Field = "UserId",
					Operator = PredicateOperator.Equals,
					Value = user.Id.ToString()
				};

				var paging = new Paging
				{
					Index = 0,
					Size = 1000
				};

				var searchAccountsRequest = new SearchAccountsRequest
				{
					Ordering = null,
					PageInfo = paging,
					Predicates = new[] { predicate }
				};

				var searchAccountsResponse =
					(await _customerManagementService.CallAsync((s, r) => s.SearchAccountsAsync(r), searchAccountsRequest));

				var accounts = searchAccountsResponse.Accounts.ToArray();
				if (accounts.Length <= 0) return;

				WriteLog("accounts.Length ::" + accounts.Length + "::");

				foreach (var account in accounts)
				{
					_authorizationData.AccountId = (long)account.Id;
					_authorizationData.CustomerId = (int)account.ParentCustomerId;
					foreach (var example in _examples)
					{
						example.RunAsyncReport(_authorizationData, System.Configuration.ConfigurationManager.AppSettings["FolderPath"].ToString()).Wait();

					}
				}




			}
			catch (Exception ex)
			{
				errorExisted = "yes";
				WriteLog("***************SetAuthorizationDataAsync Exception ::" + ex.ToString() + "::********************************");
			}

			return;
		}

		public static void WriteLog(string strLog)
		{
			StreamWriter log;
			FileStream fileStream = null;
			DirectoryInfo logDirInfo = null;
			FileInfo logFileInfo;

			string logFilePath = System.Configuration.ConfigurationManager.AppSettings["LogFolderPath"].ToString();
			logFilePath = logFilePath + "Log-File_Creation_" + System.DateTime.Today.ToString("MM-dd-yyyy") + "." + "txt";
			logFineName = logFilePath;
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
