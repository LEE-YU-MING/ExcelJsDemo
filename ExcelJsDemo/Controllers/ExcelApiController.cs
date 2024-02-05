using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Net.Http;
using System.Web.Http;

namespace ExcelJsDemo.Controllers
{
    public class ExcelApiController : ApiController
    {
        // POST: api/ExcelApi
        public string Post(IEnumerable<ProductExcel> products)
        {
            try
            {
                IEnumerable<ProductCategory> categorylist = new List<ProductCategory>()
                {
                    new ProductCategory
                    {
                        Id=1,
                        Name="分類1"
                    },
                    new ProductCategory
                    {
                        Id=2,
                        Name="分類2"
                    },
                    new ProductCategory
                    {
                        Id=3,
                        Name="分類3"
                    }
                };

                var dictionary = categorylist.ToDictionary(x => x.Name.ToLower(), x => x.Id);
                var datas = new List<ProductDto>();
                foreach(var product in products)
                {
                    var dto = new ProductDto
                    {
                        Name = product.productName,
                        DisplayOrder = product.DisplayOrder,
                        Present = product.Present,
                        UnitPrice = product.UnitPrice,
                        CategoryName = product.productCategory,
                        //用名子去找資料是否存在,存在就取得對應Id,沒存在設-1,service裡面的Create就要驗證-1 要擋住
                        CategoryId = dictionary.ContainsKey(product.productCategory.ToLower()) ? dictionary[product.productCategory.ToLower()] : -1
                    };
                    //加入資料庫
                    //service.Create(dto);
                    datas.Add(dto);
                }
                //可以看資料
                var count=datas.Count();
                
                return "true";
            }catch (Exception ex)
            {
                return ex.Message;
            }
        }

    }
    public class ProductCategory
    {
        public int Id { get; set; }
        public string Name { get; set; }
    }
    public class ProductExcel
    {
        public int UnitPrice { get; set; }
        public int DisplayOrder { get; set; }
        public string productName { get; set; }
        public string productCategory { get; set; }
        public string Present { get; set; }

    }
    public class ProductDto
    {
        public int Id { get; set; }
        public int UnitPrice { get; set; }
        public int DisplayOrder { get; set; }
        public string Name { get; set; }
        public int CategoryId { get; set; }
        public string CategoryName { get; set; }
        public string Present { get; set; }
    }
}
