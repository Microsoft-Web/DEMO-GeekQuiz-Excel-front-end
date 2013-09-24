using GeekQuiz.Models;
using GeekQuiz.Services;
using GeekQuiz.ViewModels;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Threading.Tasks;
using System.Web.Http;

namespace GeekQuiz.Controllers
{
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

        public async Task<StatisticsViewModel> Get()
        {
            StatisticsViewModel statistics =
                await this.statisticsService.GenerateStatistics();

            return statistics;
        }
    }
}
