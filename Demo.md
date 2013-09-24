<a name="title" />
# Building an Excel front end (using apps for Office) #

---
<a name="Overview" />
## Overview ##

This demo demonstrates an Excel app pulls statistics from the GeekQuiz API. 

In this demo you will:

1. Add a new empty Web API StatisticsController to the GeekQuiz application.
1. Using snippets, add a Get method which calls into StatisticsService.GenerateStatistics(). 
1. File / New / Office / Excel Task bar application.
1. Use snippet to add a call in to the GeekQuiz StatisticsController and puts the results in the Excel document.
1. Generate a quick chart or graph in Excel using the returned data.


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
1. Open the **GeekQuiz.sln** solution located under **source\begin**.
1. If you don't have one, create a user account for the application. To do that, press **F5**, click **Register** and provide the information required. After that, close the browser window.

	> **Note:** Remember the information you provided as you will be using it during the demo.

1. In Visual Studio, close all open files.
1. Make sure that you have an Internet connection, as requires the download of NuGet packages.

<a name="Demo" />
## Demo ##
This demo is composed of the following segments:

1. [Create the StatisticsController](#segment1).
1. [(TODO: Insert Segment 2 title here)](#segment2).

<a name="segment1" />
### Create the StatisticsController ###

1. Right click on the **Controllers** folder and go to **Add/Controller...** in order to create a new **StatisticsController**.

	![Creating a new Controller](images/creating-a-new-controller.png?raw=true "Creating a new Controller")

	_Creating a new Controller_

1. In the **Add Scaffold** dialog select the **Web API 2 Controller - Empty** option from the list and click **Ok**

	![Selecting the Web API 2 Controller - Empty option](images/selecting-the-web-api-controller-scaffold.png?raw=true "Selecting the Web API 2 Controller - Empty option")

	_Selecting the Web API 2 Controller - Empty option_

1. In the **Add Controller** dialog, set the Controller name to **StatisticsController**.

	![Setting the name to the StatisticsController](images/setting-the-name-to-the-statisticscontroller.png?raw=true "Setting the name to the StatisticsController")

	_Setting the name to the StatisticsController_

1. Implement the controller using the following code.

	<!-- mark:3-16 -->
	````C#
    public class StatisticsController : ApiController
    {
        private TriviaContext db;
        private StatisticsService statisticsService;

        public StatisticsController()
        {
            this.db = new TriviaContext();
            this.statisticsService = new StatisticsService(db);
        }

        protected override void Dispose(bool disposing)
        {
            this.db.Dispose();
            base.Dispose(disposing);
        }
    }
````

1. Add the following using statements.

	<!-- mark:1-2 -->
	````C#
	using GeekQuiz.Models;
	using GeekQuiz.Services;
````

1. Add the following code to create a **Get** action in the **StatisticsController**.

	<!-- mark:1-14 -->
	````C#
	public async Task<StatisticsViewModel> Get()
	{
		StatisticsViewModel statistics =
			 await this.statisticsService.GenerateStatistics();

		return statistics;
	}
````

1. Resolve the missing _using_ statements for **Task** and **StatisticsViewModel**.

	<!-- mark:1-2 -->
	````C#
	using GeekQuiz.ViewModels;
	using System.Threading.Tasks;
````


1. Build the solution.


<a name="segment2" />
### Creating an Excel Task bar application ###

1. TODO: Review and define the best flow to do the following:
	1. File / New / Office / Excel Task bar application.
	1. Use snippet to add a call in to the GeekQuiz StatisticsController and puts the results in the Excel document. 

<a name="segment3" />
### Running the excel app ###

1. TODO: Review and define the best flow to showcase the running solution
	1. Debug the application with **F5**.

		> **Note:** If the Log in page is displayed, provide the credentials you created during the setup steps.
		
		> ![Logging in the site](images/logging-in-the-app.png?raw=true "Logging in the site")

---

<a name="summary" />
## Summary ##

(TODO: Insert a summary text here. For example:)  
By completing this demo lab you have learned how to:

 * (TODO: Insert outcome 1 here)
 * (TODO: Insert outcome 2 here)
 * (TODO: Insert outcome 3 here)

---