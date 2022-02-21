using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.Rendering;
using Microsoft.EntityFrameworkCore;
using CanadaGames.Data;
using CanadaGames.Models;
using CanadaGames.Utilities;
using Microsoft.AspNetCore.Http;
using OfficeOpenXml;
using System.IO;
using Microsoft.AspNetCore.Authorization;

namespace CanadaGames.Controllers
{

    [Authorize]
    public class CoachesController : Controller
    {
        private readonly CanadaGamesContext _context;

        public CoachesController(CanadaGamesContext context)
        {
            _context = context;
        }

        // GET: Coaches
        public async Task<IActionResult> Index(string SearchString, 
            int? page, int? pageSizeID, string actionButton,
            string sortDirection = "asc", string sortField = "Last Name")
        {
            var coaches = from c in  _context.Coaches
                         select c;

            //Clear the sort/filter/paging URL Cookie for Controller
            CookieHelper.CookieSet(HttpContext, ControllerName() + "URL", "", -1);

            //Toggle the Open/Closed state of the collapse depending on if we are filtering
            ViewData["Filtering"] = ""; //Asume not filtering
            //Then in each "test" for filtering, add ViewData["Filtering"] = " show" if true;

            //NOTE: make sure this array has matching values to the column headings
            string[] sortOptions = new[] { "Last Name", "First Name" };

            if (!String.IsNullOrEmpty(SearchString))
            {
                coaches = coaches.Where(p => p.LastName.ToUpper().Contains(SearchString.ToUpper())
                                       || p.FirstName.ToUpper().Contains(SearchString.ToUpper()));
                ViewData["Filtering"] = " show";
            }

            //Before we sort, see if we have called for a change of filtering or sorting
            if (!String.IsNullOrEmpty(actionButton)) //Form Submitted!
            {
                page = 1;//Reset page to start

                if (sortOptions.Contains(actionButton))//Change of sort is requested
                {
                    if (actionButton == sortField) //Reverse order on same field
                    {
                        sortDirection = sortDirection == "asc" ? "desc" : "asc";
                    }
                    sortField = actionButton;//Sort by the button clicked
                }
            }
            //Now we know which field and direction to sort by
            if (sortField == "First Name")
            {
                if (sortDirection == "asc")
                {
                    coaches = coaches
                        .OrderBy(p => p.FirstName)
                        .ThenBy(p => p.LastName);
                }
                else
                {
                    coaches = coaches
                        .OrderByDescending(p => p.FirstName)
                        .ThenByDescending(p => p.LastName);
                }
            }
            else //Sorting by Last Name
            {
                if (sortDirection == "asc")
                {
                    coaches = coaches
                        .OrderBy(p => p.LastName)
                        .ThenBy(p => p.FirstName);
                }
                else
                {
                    coaches = coaches
                        .OrderByDescending(p => p.LastName)
                        .ThenByDescending(p => p.FirstName);
                }
            }
            //Set sort for next time
            ViewData["sortField"] = sortField;
            ViewData["sortDirection"] = sortDirection;

            //Handle Paging
            int pageSize = PageSizeHelper.SetPageSize(HttpContext, pageSizeID, ControllerName());
            ViewData["pageSizeID"] = PageSizeHelper.PageSizeList(pageSize);

            var pagedData = await PaginatedList<Coach>.CreateAsync(coaches.AsNoTracking(), page ?? 1, pageSize);

            return View(pagedData);
        }

        // GET: Coaches/Details/5
        [Authorize(Roles = "Staff,Supervisor,Admin")]
        public async Task<IActionResult> Details(int? id)
        {
            //URL with the last filter, sort and page parameters for this controller
            ViewDataReturnURL();

            if (id == null)
            {
                return NotFound();
            }

            var coach = await _context.Coaches
                .Include(c=>c.Athletes).ThenInclude(c=>c.Sport)
                .AsNoTracking()
                .FirstOrDefaultAsync(m => m.ID == id);
            if (coach == null)
            {
                return NotFound();
            }

            return View(coach);
        }

        // GET: Coaches/Create
        [Authorize(Roles = "Staff,Supervisor,Admin")]
        public IActionResult Create()
        {
            //URL with the last filter, sort and page parameters for this controller
            ViewDataReturnURL();

            return View();
        }

        // POST: Coaches/Create
        // To protect from overposting attacks, enable the specific properties you want to bind to, for 
        // more details, see http://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [Authorize(Roles = "Staff,Supervisor,Admin")]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Create([Bind("ID,FirstName,MiddleName,LastName")] Coach coach)
        {
            //URL with the last filter, sort and page parameters for this controller
            ViewDataReturnURL();

            try
            {
                if (ModelState.IsValid)
                {
                    _context.Add(coach);
                    await _context.SaveChangesAsync();
                    return RedirectToAction("Details", new { coach.ID });
                }
            }
            catch (DbUpdateException)
            {
                ModelState.AddModelError("", "Unable to save changes. Try again, and if the problem persists see your system administrator.");
            }

            return View(coach);
        }

        // GET: Coaches/Edit/5
        [Authorize(Roles = "Staff,Supervisor,Admin")]
        public async Task<IActionResult> Edit(int? id)
        {
            //URL with the last filter, sort and page parameters for this controller
            ViewDataReturnURL();

            if (id == null)
            {
                return NotFound();
            }

            var coach = await _context.Coaches.FindAsync(id);
            if (coach == null)
            {
                return NotFound();
            }

            //Staff can only edit a coach record that they created
            if (User.IsInRole("Staff"))
            {
                if (User.Identity.Name != coach.CreatedBy)
                {
                    TempData["Message"] = "You are not authorized to edit a Coach you did not enter into the system.";
                    return Redirect(ViewData["returnURL"].ToString());
                }
            }

            return View(coach);
        }

        // POST: Coaches/Edit/5
        // To protect from overposting attacks, enable the specific properties you want to bind to, for 
        // more details, see http://go.microsoft.com/fwlink/?LinkId=317598.
        [HttpPost]
        [Authorize(Roles = "Staff,Supervisor,Admin")]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> Edit(int id)
        {
            //URL with the last filter, sort and page parameters for this controller
            ViewDataReturnURL();

            //Go get the Coach to update
            var coachToUpdate = await _context.Coaches.SingleOrDefaultAsync(p => p.ID == id);

            //Check that you got it or exit with a not found error
            if (coachToUpdate == null)
            {
                return NotFound();
            }

            //Staff can only edit a coach record that they created
            if (User.IsInRole("Staff"))
            {
                if (User.Identity.Name != coachToUpdate.CreatedBy)
                {
                    TempData["Message"] = "You are not authorized to edit a Coach you did not enter into the system.";
                    return Redirect(ViewData["returnURL"].ToString());
                }
            }

            //Try updating it with the values posted
            if (await TryUpdateModelAsync<Coach>(coachToUpdate, "",
                d => d.FirstName, d => d.MiddleName, d => d.LastName))
            {
                try
                {
                    await _context.SaveChangesAsync();
                    return RedirectToAction("Details", new { coachToUpdate.ID });
                }
                catch (DbUpdateConcurrencyException)
                {
                    if (!CoachExists(coachToUpdate.ID))
                    {
                        return NotFound();
                    }
                    else
                    {
                        throw;
                    }
                }
                catch (DbUpdateException)
                {
                    ModelState.AddModelError("", "Unable to save changes. Try again, and if the problem persists see your system administrator.");
                }

            }
            return View(coachToUpdate);
        }

        // GET: Coaches/Delete/5
        [Authorize(Roles = "Supervisor,Admin")]
        public async Task<IActionResult> Delete(int? id)
        {
            //URL with the last filter, sort and page parameters for this controller
            ViewDataReturnURL();

            if (id == null)
            {
                return NotFound();
            }

            var coach = await _context.Coaches
                .AsNoTracking()
                .FirstOrDefaultAsync(m => m.ID == id);
            if (coach == null)
            {
                return NotFound();
            }

            return View(coach);
        }

        // POST: Coaches/Delete/5
        [HttpPost, ActionName("Delete")]
        [Authorize(Roles = "Supervisor,Admin")]
        [ValidateAntiForgeryToken]
        public async Task<IActionResult> DeleteConfirmed(int id)
        {
            //URL with the last filter, sort and page parameters for this controller
            ViewDataReturnURL();

            var coach = await _context.Coaches.FindAsync(id);
            try
            {
                _context.Coaches.Remove(coach);
                await _context.SaveChangesAsync();
                return Redirect(ViewData["returnURL"].ToString());
            }
            catch (DbUpdateException dex)
            {
                if (dex.GetBaseException().Message.Contains("FOREIGN KEY constraint failed"))
                {
                    ModelState.AddModelError("", "Unable to Delete Coach. Remember, you cannot delete a Coach working with Athletes.");
                }
                else
                {
                    ModelState.AddModelError("", "Unable to save changes. Try again, and if the problem persists see your system administrator.");
                }
            }
            return View(coach);

        }

        [HttpPost]
        [Authorize(Roles = "Staff,Supervisor,Admin")]
        public async Task<IActionResult> InsertFromExcel(IFormFile theExcel)
        {
            //Note: We will return to a fresh copy of the Index and display 
            //a message showing the result of the import.  We will add the message
            //into TempData so it lives long enough to get to the Index View

            //Make sure a file has been uploaded
            if (theExcel == null)
            {
                TempData["Message"] = "You must select a file before you try to upload the data.";
                return RedirectToAction(nameof(Index));
            }

            string uploadMessage = "";
            int i = 0;//Counter for inserted records
            int j = 0;//Counter for duplicates

            try
            {
                ExcelPackage excel;
                using (var memoryStream = new MemoryStream())
                {
                    await theExcel.CopyToAsync(memoryStream);
                    excel = new ExcelPackage(memoryStream);
                }
                var workSheet = excel.Workbook.Worksheets[0];
                var start = workSheet.Dimension.Start;
                var end = workSheet.Dimension.End;

                //For the bonus, pull the current list of Coaches into a HashSet of the concatenated names.
                //This is faster then querying the database over and over again to check for duplicates.
                //Note: we can't expect the HashSet to detect duplicate objects without defining how to compare them
                //so it is easier to work with a string concatenating the 3 name components together.
                var existingCoaches = (_context.Coaches
                    .Select(c => new { name = c.FirstName + c.MiddleName + c.LastName }))
                    .ToList().Select(c => c.name).ToHashSet(); 
                    //Note we first get the collection of name objects and then
                    //make a HashSet out of it.
                
                //Start a new list to hold imported objects
                List<Coach> coaches = new List<Coach>();

                //The first row contains columns headings
                for (int row = start.Row + 1; row <= end.Row; row++) 
                {
                    // Row by row...
                    Coach c = new Coach
                    {
                        FirstName = workSheet.Cells[row, 1].Text,
                        MiddleName = workSheet.Cells[row, 2].Text,
                        LastName = workSheet.Cells[row, 3].Text
                    };
                    //Check for duplicate of concatenated name
                    if (existingCoaches.Contains(c.FirstName + c.MiddleName + c.LastName))
                    {
                        j++;
                    }
                    else
                    {
                        coaches.Add(c);
                        i++;
                    }
                }
                _context.Coaches.AddRange(coaches);
                _context.SaveChanges();
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.GetBaseException().Message);
                uploadMessage = "Failed to import data.  Check that you selected the correct file in the correct format.";
            }
            if(String.IsNullOrEmpty(uploadMessage))
            {
                uploadMessage = "Imported " + (i + j).ToString() + " records, with "
                    + j.ToString() + " rejected as duplicates and " + i.ToString() + " inserted.";
            }
            TempData["Message"] = uploadMessage;
            return RedirectToAction(nameof(Index));
        }


        private string ControllerName()
        {
            return this.ControllerContext.RouteData.Values["controller"].ToString();
        }
        private void ViewDataReturnURL()
        {
            ViewData["returnURL"] = MaintainURL.ReturnURL(HttpContext, ControllerName());
        }

        private bool CoachExists(int id)
        {
            return _context.Coaches.Any(e => e.ID == id);
        }
    }
}
