(function () {
    "use strict";

    var bindingID = "statisticsTableId";
    var tableName = "StatisticsTable";

    // The initialize function must be run each time a new page is loaded
    Office.initialize = function (reason) {
        $(document).ready(function () {
            $('#update-statistics').click(updateStatisticsTable);
        });
    };

    function retriveData(callback) {
        $.getJSON("../api/statistics", function (data) {
            callback(data);
        });
    }

    function convertDataToRows(data) {
        var rows = [];
        var row = [data.totalAnswers,
                    data.correctAnswers,
                    data.incorrectAnswers,
                    data.correctAnswersAverage,
                    data.incorrectAnswersAverage,
                    data.totalAnswersAverage];

        rows.push(row);

        return rows;
    }

    // Update the TableData object referenced by the binding 
    // and then update the data in the table on the worksheet. 
    function updateStatisticsTable() {
        Office.context.document.bindings.addFromNamedItemAsync(
          tableName,
          Office.BindingType.Table,
          { id: bindingID },
          function (asyncResult) {
              if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
                  retriveData(function (data) {
                      var headers = [['Total', 'Correct', 'Incorrect', 'Correct p/user', 'Incorrect p/user', 'Total p/user']];
                      var newValuesTable = new Office.TableData(convertDataToRows(data), headers);

                      asyncResult.value.setDataAsync(newValuesTable, { coercionType: Office.CoercionType.Table });
                  });
              }
          });      
    }
})();