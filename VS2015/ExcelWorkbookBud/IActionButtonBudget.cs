using System;
namespace ExcelWorkbook.Actions
{
    interface IActionButtonDatas
    {
        void call();
        string Mesret { get; }
        int Staret { get; }
    }
}
