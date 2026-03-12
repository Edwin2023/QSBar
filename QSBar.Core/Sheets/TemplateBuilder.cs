namespace QSBar.Core.Sheets
{
    using System.Collections.Generic;

    public static class TemplateBuilder
    {
        public static string[] GetDefaultHeaders()
        {
            var order = new[] { TemplateFields.SKU, TemplateFields.Name, TemplateFields.Price, TemplateFields.TaxRate, TemplateFields.Quantity };
            var list = new List<string>(order.Length);
            foreach (var f in order)
                list.Add(MapFieldName(f));
            return list.ToArray();
        }

        public static string MapFieldName(TemplateFields f)
        {
            switch (f)
            {
                case TemplateFields.SKU: return "SKU";
                case TemplateFields.Name: return "名称";
                case TemplateFields.Price: return "价格";
                case TemplateFields.TaxRate: return "税率";
                case TemplateFields.Quantity: return "数量";
                default: return f.ToString();
            }
        }
    }
}
