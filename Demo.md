<a name="title" />
# Building an Excel front end (using apps for Office) #

---
<a name="Overview" />
## Overview ##

This demo demonstrates an Excel app pulls statistics from the GeekQuiz API. 

<a id="goals" />
### Goals ###
In this demo, you will see how to:

1. (TODO: Insert goal 1 here)
1. (TODO: Insert goal 2 here)
1. (TODO: Insert goal 3 here)

<a name="technologies" />
### Key Technologies ###

- {TODO: Include technology name here} [here][1]
- {TODO: Include technology name here}
- [{TODO: Include technology name here}][2]

[1]: http://insert_link_to_technology_1_here/
[2]: http://insert_link_to_technology_2_here/

<a name="Setup" />
### Setup and Configuration ###
Follow these steps to setup your environment for the demo.

1. Open Visual Studio 2013.
1. Open the **GeekQuiz.sln** solution located under **source\end**.
1. If you don't have one, create a user account for the application. To do that, press **F5**, click **Register** and provide the information required. After that, close the browser window.

	> **Note:** Remember the information you provided as you will be using it during the demo.

1. Answers a few questions.
1. Make sure that the **GeekQuiz Website** project has the **Current Page** configured as **Start Action**. To do this, open the project properties and open the Web tab.

	![Configuring the start action for the web site](images/configuring-the-start-action-of-the-website.png?raw=true "Configuring the start action for the web site")

	_Configuring the start action for the web site_
 
1. In Visual Studio, close all open files.
1. Make sure that you have an Internet connection, as requires the download of NuGet packages.
1. Make sure that you have **Microsoft Excel 2013** installed.

<a name="Demo" />
## Demo ##
This demo is composed of the following segments:

1. [Exploring the App for Office](#segment1).
1. [Running the solution](#segment2).

<a name="segment1" />
### Exploring the App for Office ###

1. Expand the **Controllers** folder, open the  **StatisticsController** file and show the **Get** method.

	<!-- mark:1-13 -->
	````C#
	// GET api/statistics
	[ResponseType(typeof(StatisticsViewModel))]
	public async Task<IHttpActionResult> Get()
	{
		StatisticsViewModel statistics =
			 await this.statisticsService.GenerateStatistics();
		if (statistics == null)
		{
			 return NotFound();
		}

		return Ok(statistics);
	}
````

1. In the same folder open the **OfficeAppController** file and show that the **Index** action returns a view.

	<!-- mark:5-8 -->
	````C#
	public class OfficeAppController : Controller
	{
		//
		// GET: /Office/
		public ActionResult Index()
		{
			return View();
		}
	}
````

1. Open the **GeekQuiz.OfficeManifest** located in the **GeekQuiz.Office** project and show that the **Source location** is defined as **GeekQuiz/OfficeApp/**.

	![Showing the Office Manifest](images/showing-the-office-manifest.png?raw=true "Showing the Office Manifest")

	_Showing the Office Manifest_

1. Go back to the **GeekQuiz** project and open the **Index.cshtml** file located in the **Views/OfficeApp** folder.

1. Show button in that page

	````HTML
	<button id="update-statistics" disabled >Update Statistics</button>
	````

1. Show the **Scripts** section at the end of the file.

	````HTML
	@section Scripts {
		 <script src="@Url.Content("~/Scripts/OfficeApp.js")"></script>
	}
	````

1. Open the **OfficeApp.js** file located in the **Scripts** folder

1. Show the **Office.initialize** statement.

	````JavaScript
	// The initialize function must be run each time a new page is loaded
    Office.initialize = function (reason) {
        $(document).ready(function () {
            $('#update-statistics').click(updateStatisticsTable);

            initializeBindings();
        });
    };
````

1. Show the **initializeBindings** function

	````JavaScript
	function initializeBindings() {
        Office.context.document.bindings.addFromNamedItemAsync(
          tableName,
          Office.BindingType.Table,
          { id: bindingID },
          function (asyncResult) {
              if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
                  $('#update-statistics').prop("disabled", false);
              }
          });
    }
````

1. Finally, show the **updateStatisticsTable** function

	````JavaScript
	function updateStatisticsTable() {
        $.getJSON("/api/statistics", function (data) {
            var headers = [['Total', 'Correct', 'Incorrect', 'Correct p/user', 'Incorrect p/user', 'Total p/user']];
            var rows = [[data.totalAnswers, data.correctAnswers, data.incorrectAnswers,
                          data.correctAnswersAverage, data.incorrectAnswersAverage, data.totalAnswersAverage]];
            var newValuesTable = new Office.TableData(rows, headers);

            Office.select("bindings#" + bindingID).setDataAsync(newValuesTable, { coercionType: Office.CoercionType.Table });
        });
    }
````

<a name="segment2" />
### Running the solution ###

1. Debug the application with **F5**.

	![Running the solution](images/running-the-solution.png?raw=true "Running the solution")
	
	_Running the solution_

1. Once the Excel document is open, show the app for Office.

	![Showing the app for office](images/showing-the-app-for-office.png?raw=true "Showing the app for office")
	
	_Showing the app for office_

1. Navigate to the **DESIGN** tab and show that the **Table Name** is **StatisticsTable**.

	![Showing the table name](images/showing-the-table-name.png?raw=true "Showing the table name")
	
	_Showing the table name_

1. Click the **Update Statistics** button.

	![Updating the statistics](images/updating-the-statistics.png?raw=true "Updating the statistics")
	
	_Updating the statistics_

1. Show the new data in the statistics table.

	![Showing the updated statistics](images/updated-statistics.png?raw=true "Showing the updated statistics")
	
	_Showing the updated statistics_

1. Switch to the **GeekQuiz** Web site.

	> **Note:** If the Log in page is displayed, provide the credentials you created during the setup steps.
	
	> ![Logging in the site](images/logging-in-the-app.png?raw=true "Logging in the site")	

1. Answers a few questions.

1. Go back to **Excel** and click the **update statistics** button one more time.

1. Show that the data has changed again with the latest answers.

---

<a name="summary" />
## Summary ##

(TODO: Insert a summary text here. For example:)  
By completing this demo lab you have learned how to:

 * (TODO: Insert outcome 1 here)
 * (TODO: Insert outcome 2 here)
 * (TODO: Insert outcome 3 here)

---