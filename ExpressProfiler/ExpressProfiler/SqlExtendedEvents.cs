using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExpressProfiler
{
    class SqlExtendedEvents
    {

		/*
         
		This illustrates 2 of the new/improved way to get the profile information using Extended Events. The following code shows 2 different ways to use it.

		Before running the code you need to start a session: https://docs.microsoft.com/en-us/sql/t-sql/statements/create-event-session-transact-sql?view=sql-server-ver15

		eg:

			CREATE EVENT SESSION [sample_session] ON SERVER 
			ADD EVENT sqlserver.sql_statement_completed(
			    ACTION(sqlserver.sql_text))
			GO
			
			ALTER EVENT SESSION [sample_session] ON SERVER 
				STATE = START
			GO
			
	    and 

        -- DROP EVENT SESSION [sample_session] ON SERVER





         
//Specify these two parameters.
private static string sqlInstanceName = "ThinkPad";
private static string xeSessionName = "system_health";

void Main()
{
	SqlConnectionStringBuilder csb = new SqlConnectionStringBuilder();
	csb.DataSource = sqlInstanceName;
	csb.InitialCatalog = "master";
	csb.IntegratedSecurity = true;

	OutputXELStream(csb.ConnectionString, "sample_session");
}

/*static void FullExample()
{
	try
	{
		//Connection string builder for SQL 
		//(Windows Authentication is assumed).
		SqlConnectionStringBuilder csb = new SqlConnectionStringBuilder();
		csb.DataSource = sqlInstanceName;
		csb.InitialCatalog = "master";
		csb.IntegratedSecurity = true;

		using (QueryableXEventData xEvents =
			new QueryableXEventData(
				csb.ConnectionString,
				xeSessionName,
				EventStreamSourceOptions.EventStream,
				EventStreamCacheOptions.DoNotCache))
		{
			foreach (PublishedEvent evt in xEvents)
			{
				Console.ForegroundColor = ConsoleColor.Green;
				Console.WriteLine(evt.Name);
				Console.ForegroundColor = ConsoleColor.Yellow;

				foreach (PublishedEventField fld in evt.Fields)
				{
					Console.WriteLine("\tField: {0} = {1}",
						fld.Name, fld.Value);
				}

				foreach (PublishedAction act in evt.Actions)
				{
					Console.WriteLine("\tAction: {0} = {1}",
						act.Name, act.Value);
				}

				Console.WriteLine(Environment.NewLine +
					Environment.NewLine);   //Whitespace

				//TODO: 
				//Handle the event here. 
				//(Send email, log to database/file, etc.)
				//This could be done entirely via C#.
				//Another option is to invoke a stored proc and 
				//handle the event from within SQL Server.

				//This simple example plays a "beep" 
				//when an event is received.
				System.Media.SystemSounds.Beep.Play();
			}
		}
	}
	catch (Exception ex)
	{
		Console.WriteLine(Environment.NewLine);
		Console.ForegroundColor = ConsoleColor.Magenta;
		Console.WriteLine(ex.ToString());
		Console.WriteLine(Environment.NewLine);
		Console.WriteLine("Press any key to exit.");
		Console.ReadKey(false);
	}
}* /

		static void OutputXELStream(string connectionString, string sessionName)
		{
			var cancellationTokenSource = new CancellationTokenSource();

			var xeStream = new XELiveEventStreamer(connectionString, sessionName);

			Console.WriteLine("Press any key to stop listening...");
			Task waitTask = Task.Run(() =>
			{
				Console.ReadKey();
				cancellationTokenSource.Cancel();
			});

			Task readTask = xeStream.ReadEventStream(() =>
			{
				Console.WriteLine("Connected to session");
				return Task.CompletedTask;
			},
				xevent =>
				{
					Console.WriteLine(xevent);
					Console.WriteLine("");
					return Task.CompletedTask;
				},
				cancellationTokenSource.Token);


			try
			{
				Task.WaitAny(waitTask, readTask);
			}
			catch (TaskCanceledException)
			{
			}

			if (readTask.IsFaulted)
			{
				Console.Error.WriteLine("Failed with: {0}", readTask.Exception);
			}
		}


		         
         */

	}
}
