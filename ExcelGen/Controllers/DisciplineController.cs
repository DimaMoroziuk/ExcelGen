using System.Collections.Generic;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using ExcelGen.Domain.DTOs;
using ExcelGen.Repository.Models;
using ExcelGen.Domain.Interfaces;
using System.Linq;
using System.Drawing;
using System.Net.Http;
using System.IO;
using OfficeOpenXml;
using System;

namespace ExcelGen.Controllers
{
    [Route("api/[controller]")]
    [ApiController]
    public class DisciplineController : ControllerBase
    {
        private IDisciplineManager _disciplineManager;

        public DisciplineController(IDisciplineManager disciplineManager)
        {
            _disciplineManager = disciplineManager;
        }

        // GET: api/Discipline
        [HttpGet]
        public async Task<ActionResult<IEnumerable<DisciplineDTO>>> GetDisciplineDTOs()
        {
            try
            {
                return new ActionResult<IEnumerable<DisciplineDTO>>(await _disciplineManager.GetAllDisciplines());
            }
            catch
            {
                return BadRequest();
            }
        }

        // GET: api/Discipline/excelGen
        [HttpGet("excelGen")]
        public async Task<ActionResult> DownloadExcelEPPlus()
        {
            var stream = new MemoryStream();
            ExcelPackage excelPackage = new ExcelPackage(stream);
            await _disciplineManager.FillWorkbook(excelPackage.Workbook);

            byte[] bytes = excelPackage.GetAsByteArray();
            string excelName = $"ExcelGen-{DateTime.Now.ToString("yyyyMMddHHmmssfff")}.xlsx";

            System.Net.Mime.ContentDisposition cd = new System.Net.Mime.ContentDisposition
            {
                FileName = excelName,
                Inline =  false
            };

            Response.Headers.Add("Content-Disposition", cd.ToString());
            Response.Headers.Add("X-Content-Type-Options", "nosniff");
            return File(bytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", excelName);
        }

        [HttpGet("excelGenSaved")]
        public async Task<ActionResult> SaveExcelEPPlus()
        {
            var stream = new MemoryStream();
            ExcelPackage excelPackage = new ExcelPackage(stream);
            await _disciplineManager.FillWorkbook(excelPackage.Workbook);

            byte[] bytes = excelPackage.GetAsByteArray();
            string excelName = $"ExcelGen-{DateTime.Now.ToString("yyyyMMddHHmmssfff")}.xlsx";

            string path = @"C:\ExcelDemo.xlsx";
            System.IO.File.WriteAllBytes(path, bytes);
            return Ok();
        }

        // GET: api/Discipline/5
        [HttpGet("{id}")]
        public async Task<ActionResult<DisciplineDTO>> GetDisciplineDTO(string id)
        {
            try
            {
                return new ActionResult<DisciplineDTO>( await _disciplineManager.GetDisciplineById(id));
            }
            catch 
            {
                return BadRequest();
            }
        }

        // PUT: api/Discipline/5
        [HttpPut("{id}")]
        public async Task<IActionResult> PutDisciplineDTO(string id, DisciplineDTO disciplineDTO)
        {
            try
            {
                await _disciplineManager.UpdateExistingDiscipline(id, disciplineDTO);
            }
            catch
            {
                return BadRequest();
            }

            return NoContent();
        }

        // POST: api/Discipline
        [HttpPost]
        public async Task<ActionResult<DisciplineDTO>> PostDisciplineDTO(DisciplineDTO disciplineDTO)
        {
            try 
            {
                await _disciplineManager.CreateNewDiscipline(disciplineDTO);
            }
            catch 
            {
                return BadRequest();
            }
            return CreatedAtAction("GetDisciplineDTO", new { id = disciplineDTO.Id }, disciplineDTO);
        }

        // DELETE: api/Discipline/5
        [HttpDelete("{id}")]
        public async Task<ActionResult<Discipline>> DeleteDisciplineDTO(string id)
        {
            try
            {
                await _disciplineManager.DeleteDisciplineById(id);
            }
            catch 
            {
                return BadRequest();
            }
            return NoContent();
        }
    }
}
