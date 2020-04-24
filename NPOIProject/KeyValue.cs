namespace NPOIProject
{
    /// <summary>
    /// 自定义Key Value类 Time 2017-4-6
    /// </summary>
    /// <typeparam name="TKey">Key</typeparam>
    /// <typeparam name="TValue">Value</typeparam>
    public class KeyValue<TKey, TValue>
    {
        /// <summary>
        /// 有参构造
        /// </summary>
        /// <param name="key"></param>
        /// <param name="value"></param>
        public KeyValue(TKey key, TValue value)
        {
            Key = key;
            Value = value;
        }
        /// <summary>
        /// 有参构造
        /// </summary>
        /// <param name="key"></param>
        public KeyValue(TKey key)
        {
            Key = key;
        }
        /// <summary>
        /// 无参构造
        /// </summary>
        public KeyValue()
        {
            Key = default(TKey);
            Value = default(TValue);
        }
        /// <summary>
        /// 
        /// </summary>
        public TKey Key
        {
            get; set;
        }
        /// <summary>
        /// 
        /// </summary>
        public TValue Value
        {
            get; set;
        }
    }
}
