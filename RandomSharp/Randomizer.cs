using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace RandomSharp
{

    public class Randomizer : IRandomizer
    {
        protected static Random _Random = new Random();

        protected const string _UppercaseAlpha = "ABCDEFGHIJKLMNOPQRSTUVWXYZ";
        protected const string _LowercaseAlpha = "abcdefghijklmnopqrstuvwxyz";
        protected const string _Numeric = "0123456789";

        /// <summary>
        /// Get random value from enum
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <returns></returns>
        public T Enumeration<T>() where T : struct
        {
            if (typeof(T).IsEnum)
            {
                IList<T> list = Enum.GetValues(typeof(T)).Cast<T>().ToList();
                return list[_Random.Next(list.Count)];
            }
            else
            {
                throw new InvalidCastException("Type must be enum");
            }
        }

        /// <summary>
        /// Get random value from enum
        /// returned enum can be null
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <returns></returns>
        public T? NullableEnumeration<T>() where T : struct
        {
            return Nullable<T?>(() => Enumeration<T>());
        }

        /// <summary>
        /// Get random Date between two dates
        /// </summary>
        /// <param name="min">min date</param>
        /// <param name="max">max date</param>
        /// <returns>Date</returns>
        public DateTime Date(DateTime min, DateTime max)
        {
            if (max <= min)
                return min;
            TimeSpan t_max = max - System.DateTime.MinValue;
            TimeSpan t_min = min - System.DateTime.MinValue;
            int days = _Random.Next(t_min.TotalDays.ToInt(), t_max.TotalDays.ToInt());
            return System.DateTime.MinValue.AddDays(days);
        }

        /// <summary>
        /// Get random Date between two dates
        /// returned date can be null
        /// </summary>
        /// <param name="min">min date</param>
        /// <param name="max">max date</param>
        /// <returns>date</returns>
        public DateTime? NullableDate(DateTime min, DateTime max)
        {
            return Nullable<DateTime?>(()=>Date(min, max));
        }

        /// <summary>
        /// Get random Date with time between two dates
        /// </summary>
        /// <param name="min">min datetime</param>
        /// <param name="max">max datetime</param>
        /// <returns>datetime</returns>
        public DateTime DateTime(DateTime min, DateTime max)
        {
            if (max <= min)
                return min;
            var rand_seconds = (max - min).TotalSeconds * _Random.NextDouble();
            return min.AddSeconds(rand_seconds);
        }

        /// <summary>
        /// Get random Date with time between two dates
        /// returned datetime can be null
        /// </summary>
        /// <param name="min">min datetime</param>
        /// <param name="max">max datetime</param>
        /// <returns>datetime</returns>
        public DateTime? NullableDateTime(DateTime min, DateTime max)
        {
            return Nullable<DateTime?>(() => DateTime(min, max));
        }

        /// <summary>
        /// Get random boolean
        /// </summary>
        /// <returns>bool</returns>
        public bool Boolean()
        {
            return InternalBoolean();
        }

        /// <summary>
        /// Get random boolean
        /// returned boolean can be null
        /// </summary>
        /// <returns>bool</returns>
        public bool? NullableBoolean()
        {
            return Nullable<bool?>(() => InternalBoolean());
        }

        /// <summary>
        /// Get random Int between two integers
        /// </summary>
        /// <param name="min">min int</param>
        /// <param name="max">max int</param>
        /// <returns>int</returns>
        public int Int(int min, int max)
        {
            return _Random.Next(min, max + 1);
        }

        /// <summary>
        /// Get random Int between two integers
        /// returned Int can be null
        /// </summary>
        /// <param name="min">min int</param>
        /// <param name="max">max int</param>
        /// <returns>int</returns>
        public int? NullableInt(int min, int max)
        {
            return Nullable<int?>(() => Int(min, max));
        }

        /// <summary>
        /// Get non-negative random integer that is less than or equal to the specified maximum.
        /// </summary>
        /// <param name="max">max int</param>
        /// <returns>int</returns>
        public int Int(int max)
        {
            return _Random.Next(max + 1);
        }

        /// <summary>
        /// Get random Double between two doubles
        /// </summary>
        /// <param name="min">min double</param>
        /// <param name="max">max double</param>
        /// <returns>double</returns>
        public double Double(double min, double max)
        {
            return _Random.NextDouble() * (max - min) + min;
        }

        /// <summary>
        /// Get random Double between two doubles
        /// returned Double can be null
        /// </summary>
        /// <param name="min">min double</param>
        /// <param name="max">max double</param>
        /// <returns>double</returns>
        public double? NullableDouble(double min, double max)
        {
            return Nullable<double?>(() => Double(min, max));
        }

        /// <summary>
        /// Get random decimal between two decimals
        /// </summary>
        /// <param name="min">min decimal</param>
        /// <param name="max">max decimal</param>
        /// <returns>decimal</returns>
        public decimal Decimal(decimal min, decimal max)
        {
            return _Random.NextDouble().ToDecimal() * (max - min) + min;
        }

        /// <summary>
        /// Get random decimal between two decimals
        /// returned decimal can be null
        /// </summary>
        /// <param name="min">min decimal</param>
        /// <param name="max">max decimal</param>
        /// <returns>decimal</returns>
        public decimal? NullableDecimal(decimal min, decimal max)
        {
            return Nullable<decimal?>(() => Decimal(min, max));
        }

        /// <summary>
        /// Get random value from list
        /// </summary>
        /// <typeparam name="T">list type</typeparam>
        /// <param name="list">list</param>
        /// <returns>value</returns>
        public T Random<T>(IList<T> list)
        {
            return list != null && list.Any() ? list[_Random.Next(list.Count)] : default(T);
        }

        /// <summary>
        /// Get random value from parameters
        /// </summary>
        /// <typeparam name="T">list type</typeparam>
        /// <param name="list">list</param>
        /// <returns>value</returns>
        public T Random<T>(params T[] list) where T : struct
        {
            return list != null && list.Any() ? list[_Random.Next(list.Length)] : default(T);
        }

        /// <summary>
        /// Generate random string
        /// </summary>
        /// <param name="length">length of the random string</param>
        /// <param name="stringCharacterType">specify the desired type of characters to include in the generated strings</param>
        /// <returns>string</returns>
        public string String(int length, StringCharacterType stringCharacterType = StringCharacterType.UppercaseAlpha)
        {
            switch (stringCharacterType)
            {
                case StringCharacterType.Numeric:
                    return String(length, _Numeric);
                case StringCharacterType.UppercaseAlphaNumeric:
                    return String(length, string.Concat(_UppercaseAlpha, _Numeric));
                case StringCharacterType.LowercaseAlphaNumeric:
                    return String(length, string.Concat(_LowercaseAlpha, _Numeric));
                case StringCharacterType.MixedAlphaNumeric:
                    return String(length, string.Concat(_UppercaseAlpha, _LowercaseAlpha, _Numeric));
                case StringCharacterType.UppercaseAlpha:
                    return String(length, _UppercaseAlpha);
                case StringCharacterType.LowercaseAlpha:
                    return String(length, _LowercaseAlpha);
                case StringCharacterType.MixedAlpha:
                    return String(length, string.Concat(_UppercaseAlpha, _LowercaseAlpha));
                default:
                    return String(length, _UppercaseAlpha);
            }
        }

        /// <summary>
        /// Generate random string
        /// </summary>
        /// <param name="minLength">minimum length of the random string</param>
        /// <param name="maxLength">maximum length of the random string</param>
        /// <param name="stringCharacterType">specify the desired type of characters to include in the generated strings</param>
        /// <returns>string</returns>
        public string String(int minLength, int maxLength, StringCharacterType stringCharacterType = StringCharacterType.UppercaseAlpha)
        {
            int length = Int(minLength, maxLength);
            return String(length, stringCharacterType);
        }

        /// <summary>
        /// Generate random string
        /// returned string can be null
        /// </summary>
        /// <param name="minLength">minimum length of the random string</param>
        /// <param name="maxLength">maximum length of the random string</param>
        /// <param name="stringCharacterType">specify the desired type of characters to include in the generated strings</param>
        /// <returns>string</returns>
        public string NullableString(int minLength, int maxLength, StringCharacterType stringCharacterType = StringCharacterType.UppercaseAlpha)
        {
            return Nullable<string>(() => String(minLength, maxLength, stringCharacterType));
        }

        /// <summary>
        /// Generate random string
        /// </summary>
        /// <param name="length">length of the random string</param>
        /// <param name="fromChars">specify the desired characters to include in the generated strings</param>
        /// <returns>string</returns>
        public string String(int length, string fromChars)
        {
            if (string.IsNullOrWhiteSpace(fromChars))
                throw new ArgumentException($"Invalid {nameof(fromChars)} parameter");
            var stringBuilder = new StringBuilder(length);

            for (int i = 0; i < length; i++)
            {
                char randomChar = fromChars[_Random.Next(fromChars.Length)];
                stringBuilder.Append(randomChar);
            }

            return stringBuilder.ToString();
        }

        /// <summary>
        /// Generate random string from chars
        /// </summary>
        /// <param name="minLength">minimum length of the random string</param>
        /// <param name="maxLength">maximum length of the random string</param>
        /// <param name="fromChars">specify the desired characters to include in the generated strings</param>
        /// <returns>string</returns>
        public string String(int minLength, int maxLength, string fromChars)
        {
            int length = Int(minLength, maxLength);
            return String(length, fromChars);
        }

        /// <summary>
        /// Generate random with nullable
        /// </summary>
        /// <returns>null or func result</returns>
        public TOutput Nullable<TOutput>(Func<TOutput> func)
        {
            return InternalBoolean() ? func.Invoke() : default(TOutput);
        }

        /// <summary>
        /// Generate random with nullable
        /// </summary>
        /// <returns>null or func result</returns>
        public TOutput Nullable<TInput, TOutput>(TInput input, Func<TInput, TOutput> func)
        {
            return InternalBoolean() ? func.Invoke(input) : default(TOutput);
        }

        /// <summary>
        /// Generate random with nullable
        /// </summary>
        /// <returns>null or func result</returns>
        public TOutput Nullable<TInput1, TInput2, TOutput>(TInput1 input1, TInput2 input2, Func<TInput1, TInput2, TOutput> func)
        {
            return InternalBoolean() ? func.Invoke(input1, input2) : default(TOutput);
        }

        /// <summary>
        /// Get random boolean
        /// </summary>
        /// <returns>bool</returns>
        protected virtual bool InternalBoolean()
        {
            return _Random.Next(0, 2) == 0;
        }
    }
}