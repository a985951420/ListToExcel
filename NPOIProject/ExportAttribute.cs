using System;

namespace NPOIProject
{
    /// <summary>
    /// 导出属性
    /// 
    /// 对于该标识字段，必须导出
    /// </summary>
    public class ExportAttribute : Attribute
    {
        /// <summary>
        /// 
        /// </summary>
        public ExportAttribute()
        {
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="headerName"></param>
        public ExportAttribute(string headerName)
        {
            HeaderName = headerName;
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="headerName"></param>
        /// <param name="index"></param>
        public ExportAttribute(string headerName, int index)
        {
            HeaderName = headerName;
            Index = index;
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="index"></param>
        public ExportAttribute(int index)
        {
            Index = index;
        }
        /// <summary>
        /// 
        /// </summary>
        public string HeaderName
        {
            set; get;
        }
        /// <summary>
        /// 
        /// </summary>
        public int Index
        {
            get; set;
        }
    }

    /// <summary>
    /// 不导出
    /// 
    /// 默认字段全部导出，若有显示标注不导出，则不进行导出处理
    /// </summary>
    public class NoneExportAttribute : Attribute
    {
        /// <summary>
        /// 
        /// </summary>
        public NoneExportAttribute()
        {
        }
    }
}
