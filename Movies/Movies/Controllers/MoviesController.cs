using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web.Http;

namespace Movies.Controllers
{
    public class MoviesController : ApiController
    {
        [HttpGet]
        public IHttpActionResult Demo(int id)
        {
            List<int> l = new List<int>() { 1, 2, 3 };
            return Ok(l.Where(x => x == 1).ToList());
            // return Ok(new List<int>() { 1, 2, 3 });
        }
    }
}
