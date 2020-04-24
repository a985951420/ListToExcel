using System;

namespace NPOIProject
{
    /// <summary>
    /// KVO Time 2017-5-4
    /// </summary>
    /// <typeparam name="TKey"></typeparam>
    /// <typeparam name="TValue"></typeparam>
    /// <typeparam name="TOption"></typeparam>
    [Serializable]
    public class KeyValueOption<TKey, TValue, TOption>
    {
        /// <summary>
        /// 
        /// </summary>
        public KeyValueOption()
        {
        }

        /// <summary>
        /// 
        /// </summary>
        /// <param name="key"></param>
        /// <param name="value"></param>
        /// <param name="option"></param>
        public KeyValueOption(TKey key, TValue value, TOption option)
        {
            Key = key;
            Value = value;
            Option = option;
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
        /// <summary>
        /// 
        /// </summary>
        public TOption Option
        {
            get; set;
        }
    }
}
