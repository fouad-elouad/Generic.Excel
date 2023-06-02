using System;
using System.Collections.Generic;

namespace RandomSharp
{

    public interface IRandomizer
    {

        /// <summary>
        /// Get random value from enum
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <returns></returns>
        T Enumeration<T>() where T : struct;

        /// <summary>
        /// Get random value from enum
        /// returned enum can be null
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <returns></returns>
        T? NullableEnumeration<T>() where T : struct;

        /// <summary>
        /// Get random Date between two dates
        /// </summary>
        /// <param name="min">min date</param>
        /// <param name="max">max date</param>
        /// <returns>Date</returns>
        DateTime Date(DateTime min, DateTime max);

        /// <summary>
        /// Get random Date between two dates
        /// returned date can be null
        /// </summary>
        /// <param name="min">min date</param>
        /// <param name="max">max date</param>
        /// <returns>date</returns>
        DateTime? NullableDate(DateTime min, DateTime max);

        /// <summary>
        /// Get random Date with time between two dates
        /// </summary>
        /// <param name="min">min datetime</param>
        /// <param name="max">max datetime</param>
        /// <returns>datetime</returns>
        DateTime DateTime(DateTime min, DateTime max);

        /// <summary>
        /// Get random Date with time between two dates
        /// returned datetime can be null
        /// </summary>
        /// <param name="min">min datetime</param>
        /// <param name="max">max datetime</param>
        /// <returns>datetime</returns>
        DateTime? NullableDateTime(DateTime min, DateTime max);

        /// <summary>
        /// Get random boolean
        /// </summary>
        /// <returns>bool</returns>
        bool Boolean();

        /// <summary>
        /// Get random boolean
        /// </summary>
        /// <returns>bool</returns>
        bool? NullableBoolean();

        /// <summary>
        /// Get random Int between two integers
        /// </summary>
        /// <param name="min">min int</param>
        /// <param name="max">max int</param>
        /// <returns>int</returns>
        int Int(int min, int max);

        /// <summary>
        /// Get random Int between two integers
        /// returned Int can be null
        /// </summary>
        /// <param name="min">min int</param>
        /// <param name="max">max int</param>
        /// <returns>int</returns>
        int? NullableInt(int min, int max);

        /// <summary>
        /// Get non-negative random integer that is less than or equal to the specified maximum.
        /// </summary>
        /// <param name="max">max int</param>
        /// <returns>int</returns>
        int Int(int max);

        /// <summary>
        /// Get random Double between two doubles
        /// </summary>
        /// <param name="min">min double</param>
        /// <param name="max">max double</param>
        /// <returns>double</returns>
        double Double(double min, double max);

        /// <summary>
        /// Get random Double between two doubles
        /// returned Double can be null
        /// </summary>
        /// <param name="min">min double</param>
        /// <param name="max">max double</param>
        /// <returns>double</returns>
        double? NullableDouble(double min, double max);

        /// <summary>
        /// Get random decimal between two decimals
        /// </summary>
        /// <param name="min">min decimal</param>
        /// <param name="max">max decimal</param>
        /// <returns>decimal</returns>
        decimal Decimal(decimal min, decimal max);

        /// <summary>
        /// Get random decimal between two decimals
        /// returned decimal can be null
        /// </summary>
        /// <param name="min">min decimal</param>
        /// <param name="max">max decimal</param>
        /// <returns>decimal</returns>
        decimal? NullableDecimal(decimal min, decimal max);

        /// <summary>
        /// Get random value from list
        /// </summary>
        /// <typeparam name="T">list type</typeparam>
        /// <param name="list">list</param>
        /// <returns>value</returns>
        T Random<T>(IList<T> list);

        /// <summary>
        /// Get random value from parameters
        /// </summary>
        /// <typeparam name="T">list type</typeparam>
        /// <param name="list">list</param>
        /// <returns>value</returns>
        T Random<T>(params T[] list) where T : struct;

        /// <summary>
        /// Generate random string
        /// </summary>
        /// <param name="length">length of the random string</param>
        /// <param name="stringCharacterType">specify the desired type of characters to include in the generated strings</param>
        /// <returns>string</returns>
        string String(int length, StringCharacterType stringCharacterType = StringCharacterType.UppercaseAlpha);

        /// <summary>
        /// Generate random string
        /// </summary>
        /// <param name="length">length of the random string</param>
        /// <param name="fromChars">specify the desired characters to include in the generated strings</param>
        /// <returns>string</returns>
        string String(int length, string fromChars);

        /// <summary>
        /// Generate random string
        /// </summary>
        /// <param name="minLength">minimum length of the random string</param>
        /// <param name="maxLength">maximum length of the random string</param>
        /// <param name="stringCharacterType">specify the desired type of characters to include in the generated strings</param>
        /// <returns>string</returns>
        string String(int minLength, int maxLength, StringCharacterType stringCharacterType = StringCharacterType.UppercaseAlpha);

        /// <summary>
        /// Generate random string
        /// returned string can be null
        /// </summary>
        /// <param name="minLength">minimum length of the random string</param>
        /// <param name="maxLength">maximum length of the random string</param>
        /// <param name="stringCharacterType">specify the desired type of characters to include in the generated strings</param>
        /// <returns>string</returns>
        string NullableString(int minLength, int maxLength, StringCharacterType stringCharacterType = StringCharacterType.UppercaseAlpha);

        /// <summary>
        /// Generate random string from chars
        /// </summary>
        /// <param name="minLength">minimum length of the random string</param>
        /// <param name="maxLength">maximum length of the random string</param>
        /// <param name="fromChars">specify the desired characters to include in the generated strings</param>
        /// <returns>string</returns>
        string String(int minLength, int maxLength, string fromChars);

        /// <summary>
        /// Generate random with nullable
        /// </summary>
        /// <returns>null or func result</returns>
        TOutput Nullable<TOutput>(Func<TOutput> func);

        /// <summary>
        /// Generate random with nullable
        /// </summary>
        /// <returns>null or func result</returns>
        TOutput Nullable<TInput, TOutput>(TInput input, Func<TInput, TOutput> func);

        /// <summary>
        /// Generate random with nullable
        /// </summary>
        /// <returns>null or func result</returns>
        TOutput Nullable<TInput1, TInput2, TOutput>(TInput1 input1, TInput2 input2, Func<TInput1, TInput2, TOutput> func);
    }
}