namespace NPOIProject
{
    /// <summary>
    /// 各层之间交流返回状态以及信息传递
    /// </summary>
    public class ResultStateDto
    {
        /// <summary>
        /// 
        /// </summary>
        public ResultStateDto()
        {

        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="success"></param>
        public ResultStateDto(bool success)
        {
            Success = success;
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="success"></param>
        /// <param name="message"></param>
        public ResultStateDto(bool success, string message)
        {
            Success = success;
            Message = message;
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="success"></param>
        /// <param name="message"></param>
        /// <param name="oValue"></param>
        public ResultStateDto(bool success, string message, object oValue)
        {
            Success = success;
            Message = message;
            OValue = oValue;
        }
        /// <summary>
        /// 返回状态
        /// </summary>
        public bool Success
        {
            get; set;
        }
        /// <summary>
        /// 消息内容
        /// </summary>
        public string Message
        {
            get; set;
        }
        /// <summary>
        /// 其他
        /// </summary>
        public object OValue
        {
            get; set;
        }

        /// <summary>
        /// 成功数量
        /// </summary>
        public int SuccessCount
        {
            get; set;
        }
        /// <summary>
        /// 失败数量
        /// </summary>
        public int FailureCount
        {
            get; set;
        }
    }

    /// <summary>
    /// 各层之间交流返回状态以及信息传递
    /// </summary>
    public class ResultStateDto<T>
    {
        /// <summary>
        /// 
        /// </summary>
        public ResultStateDto()
        {

        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="success"></param>
        public ResultStateDto(bool success)
        {
            Success = success;
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="success"></param>
        /// <param name="message"></param>
        public ResultStateDto(bool success, string message)
        {
            Success = success;
            Message = message;
        }
        /// <summary>
        /// 
        /// </summary>
        /// <param name="success"></param>
        /// <param name="message"></param>
        /// <param name="oValue"></param>
        public ResultStateDto(bool success, string message, T oValue)
        {
            Success = success;
            Message = message;
            OValue = oValue;
        }
        /// <summary>
        /// 返回状态
        /// </summary>
        public bool Success
        {
            get; set;
        }
        /// <summary>
        /// 消息内容
        /// </summary>
        public string Message
        {
            get; set;
        }
        /// <summary>
        /// 其他
        /// </summary>
        public T OValue
        {
            get; set;
        }
    }
}
