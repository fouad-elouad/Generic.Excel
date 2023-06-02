using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace RandomSharp
{

    public class Randomizer : IRandomizer
    {
        protected const string _Chars = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789";
        protected static Random _Random = new Random();

        /// <summary>
        /// Get random value from enum
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <returns></returns>
        public T RandomEnum<T>() where T : struct
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
        /// Get random Date between two dates
        /// </summary>
        /// <param name="min">min date</param>
        /// <param name="max">max date</param>
        /// <returns>Date</returns>
        public DateTime RandomDate(DateTime min, DateTime max)
        {
            if (max <= min)
                return min;
            TimeSpan t_max = max - DateTime.MinValue;
            TimeSpan t_min = min - DateTime.MinValue;
            int days = _Random.Next(t_min.TotalDays.ToInt(), t_max.TotalDays.ToInt());
            return DateTime.MinValue.AddDays(days);
        }

        /// <summary>
        /// Get random Date between two dates
        /// returned date can be null
        /// </summary>
        /// <param name="min">min date</param>
        /// <param name="max">max date</param>
        /// <returns>date</returns>
        public DateTime? RandomNullableDate(DateTime min, DateTime max)
        {
            return _Random.Next(0, 2) == 0 ? RandomDate(min, max) : default(DateTime?);
        }

        /// <summary>
        /// Get random Date with time between two dates
        /// </summary>
        /// <param name="min">min datetime</param>
        /// <param name="max">max datetime</param>
        /// <returns>datetime</returns>
        public DateTime RandomDateTime(DateTime min, DateTime max)
        {
            if (max <= min)
                return min;
            var rand_seconds = (max - min).TotalSeconds * _Random.NextDouble();
            return min.AddSeconds(rand_seconds);
        }

        /// <summary>
        /// Get random boolean
        /// </summary>
        /// <returns>bool</returns>
        public bool RandomBool()
        {
            return _Random.Next(0, 2) == 0;
        }

        /// <summary>
        /// Get random boolean
        /// </summary>
        /// <returns>bool</returns>
        public bool? RandomNullableBool()
        {
            int rand = _Random.Next(0, 3);
            switch (rand)
            {
                case 0:
                    return false;
                case 1:
                    return true;
                default:
                    return null;
            }
        }

        /// <summary>
        /// Get random Int between two integers
        /// </summary>
        /// <param name="min">min int</param>
        /// <param name="max">max int</param>
        /// <returns>int</returns>
        public int Random(int min, int max)
        {
            return _Random.Next(min, max + 1);
        }

        /// <summary>
        /// Get non-negative random integer that is less than or equal to the specified maximum.
        /// </summary>
        /// <param name="max">max int</param>
        /// <returns>int</returns>
        public int Random(int max)
        {
            return _Random.Next(max + 1);
        }

        /// <summary>
        /// Get random Double between two doubles
        /// </summary>
        /// <param name="min">min double</param>
        /// <param name="max">max double</param>
        /// <returns>double</returns>
        public double Random(double min, double max)
        {
            return _Random.NextDouble() * (max - min) + min;
        }

        /// <summary>
        /// Get random decimal between two decimals
        /// </summary>
        /// <param name="min">min decimal</param>
        /// <param name="max">max decimal</param>
        /// <returns>decimal</returns>
        public decimal Random(decimal min, decimal max)
        {
            return _Random.NextDouble().ToDecimal() * (max - min) + min;
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
        /// <returns>string</returns>
        public string RandomString(int length)
        {

            var stringBuilder = new StringBuilder(length);

            for (int i = 0; i < length; i++)
            {
                char randomChar = _Chars[_Random.Next(_Chars.Length)];
                stringBuilder.Append(randomChar);
            }

            return stringBuilder.ToString();
        }

        /// <summary>
        /// Generate random string
        /// </summary>
        /// <param name="minLength">minimum length of the random string</param>
        /// <param name="maxLength">maximum length of the random string</param>
        /// <returns>string</returns>
        public string RandomString(int minLength, int maxLength)
        {
            int length = Random(minLength, maxLength);
            return RandomString(length);
        }

        /// <summary>
        /// Generate random with nullable
        /// </summary>
        /// <returns>null or func result</returns>
        public TOutput RandomNullable<TOutput>(Func<TOutput> func)
        {
            return _Random.Next(0, 2) == 0 ? func.Invoke() : default(TOutput);
        }

        /// <summary>
        /// Generate random with nullable
        /// </summary>
        /// <returns>null or func result</returns>
        public TOutput RandomNullable<TInput, TOutput>(TInput input, Func<TInput, TOutput> func)
        {
            return _Random.Next(0, 2) == 0 ? func.Invoke(input) : default(TOutput);
        }

        /// <summary>
        /// Generate random with nullable
        /// </summary>
        /// <returns>null or func result</returns>
        public TOutput RandomNullable<TInput1, TInput2, TOutput>(TInput1 input1, TInput2 input2, Func<TInput1, TInput2, TOutput> func)
        {
            return _Random.Next(0, 2) == 0 ? func.Invoke(input1, input2) : default(TOutput);
        }
    }
}